#!/usr/bin/env python3
"""Validate a sci-diagram-pptx scene plan and its locked source identity.

This validator has no third-party dependencies.  It applies the bundled JSON
Schema plus semantic checks that JSON Schema cannot express conveniently:
unique IDs/z-order, in-canvas normalized boxes, valid graph references,
confidence policy, explicit degradation approval, and source-manifest matching.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Sequence


VALIDATOR_VERSION = "1.0"
TOOL_NAME = "sci-diagram-pptx/validate_scene_plan.py"
EXPECTED_MANIFEST_KIND = "sci-diagram-pptx-source-manifest"
EXPECTED_PLAN_KIND = "sci-diagram-pptx-scene-plan"
EXPECTED_PANEL_MANIFEST_KIND = "sci-diagram-pptx-panel-crop-manifest"
DEGRADED_MODES = {"native-approximation", "isolated-raster"}
UNSUPPORTED_MODE = "unsupported"


def _pointer(parts: Sequence[Any]) -> str:
    if not parts:
        return ""
    return "/" + "/".join(str(part).replace("~", "~0").replace("/", "~1") for part in parts)


def _issue(code: str, path: Sequence[Any], message: str, **details: Any) -> dict[str, Any]:
    issue: dict[str, Any] = {"id": code, "code": code, "path": _pointer(path), "message": message}
    issue.update(details)
    return issue


def _is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(float(value))


def _matches_type(value: Any, expected: str) -> bool:
    checks = {
        "object": lambda item: isinstance(item, dict),
        "array": lambda item: isinstance(item, list),
        "string": lambda item: isinstance(item, str),
        "integer": lambda item: isinstance(item, int) and not isinstance(item, bool),
        "number": _is_number,
        "boolean": lambda item: isinstance(item, bool),
        "null": lambda item: item is None,
    }
    return expected in checks and checks[expected](value)


class LightweightSchemaValidator:
    """Small Draft 2020-12 subset covering the bundled scene-plan schema."""

    def __init__(self, root_schema: dict[str, Any]) -> None:
        self.root = root_schema

    def validate(self, instance: Any) -> list[dict[str, Any]]:
        errors: list[dict[str, Any]] = []
        self._visit(instance, self.root, (), errors)
        return errors

    def _resolve_ref(self, reference: str) -> dict[str, Any]:
        if not reference.startswith("#/"):
            raise ValueError(f"Only local JSON Pointer references are supported: {reference}")
        current: Any = self.root
        for token in reference[2:].split("/"):
            token = token.replace("~1", "/").replace("~0", "~")
            if not isinstance(current, dict) or token not in current:
                raise ValueError(f"Unresolvable schema reference: {reference}")
            current = current[token]
        if not isinstance(current, dict):
            raise ValueError(f"Schema reference is not an object: {reference}")
        return current

    def _visit(self, value: Any, schema: dict[str, Any], path: tuple[Any, ...], errors: list[dict[str, Any]]) -> None:
        if "$ref" in schema:
            try:
                referenced = self._resolve_ref(schema["$ref"])
            except ValueError as exc:
                errors.append(_issue("schema.invalid_ref", path, str(exc)))
                return
            self._visit(value, referenced, path, errors)
            # Draft 2020-12 permits siblings of $ref; validate them too.
            siblings = {key: item for key, item in schema.items() if key != "$ref"}
            if siblings:
                self._visit(value, siblings, path, errors)
            return

        if "const" in schema and value != schema["const"]:
            errors.append(_issue("schema.const", path, f"Expected constant value {schema['const']!r}", actual=value))
        if "enum" in schema and value not in schema["enum"]:
            errors.append(_issue("schema.enum", path, f"Value must be one of {schema['enum']!r}", actual=value))

        expected_type = schema.get("type")
        if expected_type is not None:
            accepted = expected_type if isinstance(expected_type, list) else [expected_type]
            if not any(_matches_type(value, item) for item in accepted):
                errors.append(_issue("schema.type", path, f"Expected type {' or '.join(accepted)}", actual_type=type(value).__name__))
                return

        if isinstance(value, dict):
            required = schema.get("required", [])
            for key in required:
                if key not in value:
                    errors.append(_issue("schema.required", (*path, key), f"Required property {key!r} is missing"))
            properties = schema.get("properties", {})
            for key, item in value.items():
                if key in properties:
                    self._visit(item, properties[key], (*path, key), errors)
                elif schema.get("additionalProperties") is False:
                    errors.append(_issue("schema.additional_property", (*path, key), f"Unexpected property {key!r}"))
                elif isinstance(schema.get("additionalProperties"), dict):
                    self._visit(item, schema["additionalProperties"], (*path, key), errors)

        if isinstance(value, list):
            if "minItems" in schema and len(value) < schema["minItems"]:
                errors.append(_issue("schema.min_items", path, f"Expected at least {schema['minItems']} item(s)", actual=len(value)))
            if "maxItems" in schema and len(value) > schema["maxItems"]:
                errors.append(_issue("schema.max_items", path, f"Expected at most {schema['maxItems']} item(s)", actual=len(value)))
            if schema.get("uniqueItems"):
                for index, item in enumerate(value):
                    if item in value[:index]:
                        errors.append(_issue("schema.unique_items", (*path, index), "Array items must be unique"))
            items = schema.get("items")
            if isinstance(items, dict):
                for index, item in enumerate(value):
                    self._visit(item, items, (*path, index), errors)

        if isinstance(value, str):
            if "minLength" in schema and len(value) < schema["minLength"]:
                errors.append(_issue("schema.min_length", path, f"String must contain at least {schema['minLength']} character(s)"))
            if "maxLength" in schema and len(value) > schema["maxLength"]:
                errors.append(_issue("schema.max_length", path, f"String must contain at most {schema['maxLength']} character(s)"))
            if "pattern" in schema:
                try:
                    matched = re.search(schema["pattern"], value) is not None
                except re.error as exc:
                    errors.append(_issue("schema.invalid_pattern", path, f"Invalid pattern in schema: {exc}"))
                else:
                    if not matched:
                        errors.append(_issue("schema.pattern", path, f"String does not match {schema['pattern']!r}"))

        if _is_number(value):
            if "minimum" in schema and value < schema["minimum"]:
                errors.append(_issue("schema.minimum", path, f"Value must be at least {schema['minimum']}", actual=value))
            if "maximum" in schema and value > schema["maximum"]:
                errors.append(_issue("schema.maximum", path, f"Value must be at most {schema['maximum']}", actual=value))
            if "exclusiveMinimum" in schema and value <= schema["exclusiveMinimum"]:
                errors.append(_issue("schema.exclusive_minimum", path, f"Value must be greater than {schema['exclusiveMinimum']}", actual=value))
            if "exclusiveMaximum" in schema and value >= schema["exclusiveMaximum"]:
                errors.append(_issue("schema.exclusive_maximum", path, f"Value must be less than {schema['exclusiveMaximum']}", actual=value))


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _read_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _is_valid_sha256(value: Any) -> bool:
    return isinstance(value, str) and re.fullmatch(r"[0-9a-fA-F]{64}", value) is not None


def _confidence_issue(
    value: Any,
    path: tuple[Any, ...],
    label: str,
    minimum: float,
    low_as_error: bool,
    errors: list[dict[str, Any]],
    warnings: list[dict[str, Any]],
) -> None:
    if _is_number(value) and 0 <= value <= 1 and value < minimum:
        issue = _issue(
            "confidence.below_threshold",
            path,
            f"{label} confidence {value:.3f} is below the configured {minimum:.3f} threshold",
            confidence=value,
            threshold=minimum,
        )
        (errors if low_as_error else warnings).append(issue)


def validate_semantics(
    plan: dict[str, Any],
    minimum_confidence: float,
    low_confidence_as_error: bool,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, int]]:
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    objects = plan.get("objects") if isinstance(plan.get("objects"), list) else []
    connections = plan.get("connections") if isinstance(plan.get("connections"), list) else []
    approval_root = plan.get("approvals") if isinstance(plan.get("approvals"), dict) else {}
    degradations = approval_root.get("degradations") if isinstance(approval_root.get("degradations"), list) else []

    if plan.get("kind") != EXPECTED_PLAN_KIND:
        errors.append(
            _issue(
                "scene_plan.kind",
                ("kind",),
                f"Scene plan kind must be {EXPECTED_PLAN_KIND!r}",
                actual=plan.get("kind"),
            )
        )

    # A materialized panel crop is itself the locked source.  These fields record
    # the asserted lineage and exact parent-display bounds; JSON metadata alone
    # does not prove that the crop pixels were derived from that parent.
    plan_source = plan.get("source") if isinstance(plan.get("source"), dict) else {}
    panel_fields = (
        "parent_sha256",
        "parent_width_px",
        "parent_height_px",
        "panel_manifest_sha256",
        "panel_bbox",
        "panel_label",
        "panel_selection",
    )
    if any(field in plan_source for field in panel_fields):
        if not _is_valid_sha256(plan_source.get("parent_sha256")):
            errors.append(
                _issue(
                    "source.panel_parent_hash_required",
                    ("source", "parent_sha256"),
                    "Any parent/panel metadata requires a valid parent_sha256",
                )
            )
        parent_width = plan_source.get("parent_width_px")
        parent_height = plan_source.get("parent_height_px")
        if not isinstance(parent_width, int) or isinstance(parent_width, bool) or parent_width <= 0:
            errors.append(
                _issue(
                    "source.panel_parent_width_required",
                    ("source", "parent_width_px"),
                    "Any parent/panel metadata requires a positive parent_width_px",
                )
            )
        if not isinstance(parent_height, int) or isinstance(parent_height, bool) or parent_height <= 0:
            errors.append(
                _issue(
                    "source.panel_parent_height_required",
                    ("source", "parent_height_px"),
                    "Any parent/panel metadata requires a positive parent_height_px",
                )
            )
        if not _is_valid_sha256(plan_source.get("panel_manifest_sha256")):
            errors.append(
                _issue(
                    "source.panel_manifest_hash_required",
                    ("source", "panel_manifest_sha256"),
                    "Any parent/panel metadata requires a valid panel_manifest_sha256",
                )
            )
        if not isinstance(plan_source.get("panel_label"), str) or not plan_source.get("panel_label"):
            errors.append(
                _issue(
                    "source.panel_label_required",
                    ("source", "panel_label"),
                    "Any parent/panel metadata requires a non-empty panel_label",
                )
            )
        panel_selection = plan_source.get("panel_selection")
        if not isinstance(panel_selection, dict):
            errors.append(
                _issue(
                    "source.panel_selection_required",
                    ("source", "panel_selection"),
                    "Any parent/panel metadata requires an explicit user panel_selection record",
                )
            )
        else:
            if panel_selection.get("selection_source") != "user-explicit":
                errors.append(
                    _issue(
                        "source.panel_selection_source",
                        ("source", "panel_selection", "selection_source"),
                        "Panel selection must come from an explicit user choice",
                    )
                )
            for selection_field in ("selected_by", "selected_at", "evidence"):
                value = panel_selection.get(selection_field)
                if not isinstance(value, str) or not value:
                    errors.append(
                        _issue(
                            "source.panel_selection_incomplete",
                            ("source", "panel_selection", selection_field),
                            f"panel_selection.{selection_field} must be a non-empty string",
                        )
                    )
        panel_bbox = plan_source.get("panel_bbox")
        if not isinstance(panel_bbox, dict):
            errors.append(
                _issue(
                    "source.panel_bbox_required",
                    ("source", "panel_bbox"),
                    "Any parent/panel metadata requires panel_bbox in parent-image pixels",
                )
            )
        else:
            values = tuple(panel_bbox.get(key) for key in ("x", "y", "width", "height"))
            if all(isinstance(value, int) and not isinstance(value, bool) for value in values):
                panel_x, panel_y, panel_width, panel_height = values
                source_width, source_height = plan_source.get("width_px"), plan_source.get("height_px")
                if isinstance(parent_width, int) and not isinstance(parent_width, bool) and panel_x + panel_width > parent_width:
                    errors.append(
                        _issue(
                            "source.panel_bbox_out_of_parent_x",
                            ("source", "panel_bbox"),
                            "panel_bbox extends beyond parent_width_px",
                            right=panel_x + panel_width,
                            parent_width_px=parent_width,
                        )
                    )
                if isinstance(parent_height, int) and not isinstance(parent_height, bool) and panel_y + panel_height > parent_height:
                    errors.append(
                        _issue(
                            "source.panel_bbox_out_of_parent_y",
                            ("source", "panel_bbox"),
                            "panel_bbox extends beyond parent_height_px",
                            bottom=panel_y + panel_height,
                            parent_height_px=parent_height,
                        )
                    )
                if isinstance(source_width, int) and isinstance(source_height, int) and (panel_width != source_width or panel_height != source_height):
                    errors.append(
                        _issue(
                            "source.panel_crop_size_mismatch",
                            ("source", "panel_bbox"),
                            "Locked source dimensions must equal the exact panel crop dimensions (no rescaling)",
                            panel_size=[panel_width, panel_height],
                            source_size=[source_width, source_height],
                        )
                    )

    object_by_id: dict[str, tuple[int, dict[str, Any]]] = {}
    all_entity_ids: dict[str, tuple[str, int]] = {}
    z_orders: dict[int, tuple[str, str, int]] = {}

    for index, obj in enumerate(objects):
        if not isinstance(obj, dict):
            continue
        object_id = obj.get("id")
        if isinstance(object_id, str):
            if object_id in object_by_id:
                first_index = object_by_id[object_id][0]
                errors.append(_issue("object.id_duplicate", ("objects", index, "id"), f"Object ID {object_id!r} duplicates objects[{first_index}]", duplicate_of=f"/objects/{first_index}/id"))
            else:
                object_by_id[object_id] = (index, obj)
            if object_id in all_entity_ids:
                first_kind, first_index = all_entity_ids[object_id]
                errors.append(_issue("entity.id_duplicate", ("objects", index, "id"), f"Entity ID {object_id!r} already belongs to {first_kind}[{first_index}]"))
            else:
                all_entity_ids[object_id] = ("objects", index)
        for bbox_name, space_name in (("source_bbox", "source image"), ("bbox", "slide canvas")):
            bbox = obj.get(bbox_name)
            if isinstance(bbox, dict):
                x, y, width, height = bbox.get("x"), bbox.get("y"), bbox.get("width"), bbox.get("height")
                if all(_is_number(item) for item in (x, y, width, height)):
                    if x + width > 1 + 1e-9:
                        errors.append(_issue("bbox.out_of_bounds_x", ("objects", index, bbox_name), f"Normalized {bbox_name} extends beyond the right edge of the {space_name}", x=x, width=width, right=x + width))
                    if y + height > 1 + 1e-9:
                        errors.append(_issue("bbox.out_of_bounds_y", ("objects", index, bbox_name), f"Normalized {bbox_name} extends beyond the bottom edge of the {space_name}", y=y, height=height, bottom=y + height))
        content = obj.get("content")
        if not isinstance(content, dict) or not isinstance(content.get("text"), str):
            errors.append(
                _issue(
                    "object.content_contract",
                    ("objects", index, "content"),
                    "Object content must be an object with a text string; use an empty string for objects without text",
                )
            )
        object_style = obj.get("style")
        if not isinstance(object_style, dict) or not isinstance(object_style.get("shape_geometry"), str) or not object_style.get("shape_geometry"):
            errors.append(
                _issue(
                    "object.style_contract",
                    ("objects", index, "style"),
                    "Object style must contain a non-empty shape_geometry string",
                )
            )
        elif any(name in object_style and not isinstance(object_style[name], dict) for name in ("fill", "line", "text")):
            errors.append(
                _issue(
                    "object.style_contract",
                    ("objects", index, "style"),
                    "Object style fill, line, and text values must be objects when present",
                )
            )
        z_order = obj.get("z_order")
        if isinstance(z_order, int) and not isinstance(z_order, bool):
            if z_order in z_orders:
                first_kind, first_id, first_index = z_orders[z_order]
                errors.append(_issue("z_order.duplicate", ("objects", index, "z_order"), f"z_order {z_order} is already used by {first_kind} {first_id!r} at index {first_index}"))
            else:
                z_orders[z_order] = ("object", object_id if isinstance(object_id, str) else "?", index)
        _confidence_issue(obj.get("confidence"), ("objects", index, "confidence"), f"Object {object_id!r}", minimum_confidence, low_confidence_as_error, errors, warnings)

    for index, connection in enumerate(connections):
        if not isinstance(connection, dict):
            continue
        connection_id = connection.get("id")
        if isinstance(connection_id, str):
            if connection_id in all_entity_ids:
                first_kind, first_index = all_entity_ids[connection_id]
                errors.append(_issue("entity.id_duplicate", ("connections", index, "id"), f"Connection ID {connection_id!r} already belongs to {first_kind}[{first_index}]"))
            else:
                all_entity_ids[connection_id] = ("connections", index)
        for endpoint_name in ("from", "to"):
            endpoint = connection.get(endpoint_name)
            if not isinstance(endpoint, dict):
                continue
            reference = endpoint.get("object_id")
            if isinstance(reference, str) and reference not in object_by_id:
                errors.append(_issue("connection.object_missing", ("connections", index, endpoint_name, "object_id"), f"Connection {connection_id!r} references unknown object {reference!r}", object_id=reference))
        start = connection.get("from", {}).get("object_id") if isinstance(connection.get("from"), dict) else None
        end = connection.get("to", {}).get("object_id") if isinstance(connection.get("to"), dict) else None
        if isinstance(start, str) and start == end:
            warnings.append(_issue("connection.self_loop", ("connections", index), f"Connection {connection_id!r} is a self-loop on {start!r}"))
        label_reference = connection.get("label_object_id")
        if isinstance(label_reference, str) and label_reference not in object_by_id:
            errors.append(_issue("connection.label_missing", ("connections", index, "label_object_id"), f"Connection label references unknown object {label_reference!r}"))
        connection_style = connection.get("style")
        required_style = ("kind", "arrow_at", "dash", "line_width_px", "line_color")
        style_complete = isinstance(connection_style, dict) and all(field in connection_style for field in required_style)
        if style_complete:
            style_complete = (
                isinstance(connection_style.get("kind"), str)
                and bool(connection_style.get("kind"))
                and connection_style.get("arrow_at") in {"none", "start", "end", "both"}
                and isinstance(connection_style.get("dash"), str)
                and bool(connection_style.get("dash"))
                and _is_number(connection_style.get("line_width_px"))
                and connection_style["line_width_px"] > 0
                and isinstance(connection_style.get("line_color"), str)
                and bool(connection_style.get("line_color"))
            )
        if not style_complete:
            errors.append(
                _issue(
                    "connection.style_contract",
                    ("connections", index, "style"),
                    "Connection style must provide kind, arrow_at, dash, positive line_width_px, and line_color",
                )
            )
        z_order = connection.get("z_order")
        if isinstance(z_order, int) and not isinstance(z_order, bool):
            if z_order in z_orders:
                first_kind, first_id, first_index = z_orders[z_order]
                errors.append(_issue("z_order.duplicate", ("connections", index, "z_order"), f"z_order {z_order} is already used by {first_kind} {first_id!r} at index {first_index}"))
            else:
                z_orders[z_order] = ("connection", connection_id if isinstance(connection_id, str) else "?", index)
        _confidence_issue(connection.get("confidence"), ("connections", index, "confidence"), f"Connection {connection_id!r}", minimum_confidence, low_confidence_as_error, errors, warnings)

    # Group/child references are semantic object references too.
    for index, obj in enumerate(objects):
        if not isinstance(obj, dict) or not isinstance(obj.get("children"), list):
            continue
        seen_children: set[str] = set()
        for child_index, child_id in enumerate(obj["children"]):
            if not isinstance(child_id, str):
                continue
            if child_id in seen_children:
                errors.append(_issue("object.child_duplicate", ("objects", index, "children", child_index), f"Child object {child_id!r} appears more than once"))
            seen_children.add(child_id)
            if child_id not in object_by_id:
                errors.append(_issue("object.child_missing", ("objects", index, "children", child_index), f"Group references unknown child object {child_id!r}"))
            if child_id == obj.get("id"):
                errors.append(_issue("object.child_self", ("objects", index, "children", child_index), "An object cannot contain itself"))

    degradation_by_id: dict[str, tuple[int, dict[str, Any]]] = {}
    for index, approval in enumerate(degradations):
        if not isinstance(approval, dict):
            continue
        approval_id = approval.get("id")
        if isinstance(approval_id, str):
            if approval_id in degradation_by_id:
                first_index = degradation_by_id[approval_id][0]
                errors.append(_issue("degradation.id_duplicate", ("approvals", "degradations", index, "id"), f"Degradation ID {approval_id!r} duplicates approvals.degradations[{first_index}]"))
            else:
                degradation_by_id[approval_id] = (index, approval)
        entity_ids = approval.get("entity_ids") if isinstance(approval.get("entity_ids"), list) else []
        seen_approved_entities: set[str] = set()
        for ref_index, entity_id in enumerate(entity_ids):
            if not isinstance(entity_id, str):
                continue
            if entity_id in seen_approved_entities:
                errors.append(_issue("degradation.entity_duplicate", ("approvals", "degradations", index, "entity_ids", ref_index), f"Scene entity {entity_id!r} appears twice in one degradation record"))
            seen_approved_entities.add(entity_id)
            if entity_id not in all_entity_ids:
                errors.append(_issue("degradation.entity_missing", ("approvals", "degradations", index, "entity_ids", ref_index), f"Degradation approval references unknown scene entity {entity_id!r}"))
        if approval.get("approved") is not True:
            errors.append(_issue("degradation.not_approved", ("approvals", "degradations", index, "approved"), f"Degradation record {approval_id!r} must be explicitly approved"))
        if approval.get("approval_source") != "user-explicit":
            errors.append(_issue("degradation.approval_source", ("approvals", "degradations", index, "approval_source"), f"Degradation record {approval_id!r} must cite explicit user approval"))

    degraded_count = 0
    degraded_object_count = 0
    degraded_connection_count = 0
    unsupported_count = 0
    unsupported_object_count = 0
    unsupported_connection_count = 0
    used_degradation_ids: set[str] = set()
    scene_entities = [("objects", index, item) for index, item in enumerate(objects)] + [
        ("connections", index, item) for index, item in enumerate(connections)
    ]
    for collection, index, obj in scene_entities:
        if not isinstance(obj, dict):
            continue
        object_id = obj.get("id")
        reconstruction = obj.get("reconstruction")
        if not isinstance(reconstruction, dict):
            continue
        mode = reconstruction.get("mode")
        degradation_id = reconstruction.get("degradation_id")
        if mode == UNSUPPORTED_MODE:
            unsupported_count += 1
            if collection == "connections":
                unsupported_connection_count += 1
            else:
                unsupported_object_count += 1
            errors.append(_issue("reconstruction.unsupported", (collection, index, "reconstruction", "mode"), f"Scene entity {object_id!r} is marked unsupported and blocks reconstruction"))
            continue
        if mode in DEGRADED_MODES:
            degraded_count += 1
            if collection == "connections":
                degraded_connection_count += 1
            else:
                degraded_object_count += 1
            if not isinstance(degradation_id, str):
                errors.append(_issue("degradation.approval_required", (collection, index, "reconstruction", "degradation_id"), f"Scene entity {object_id!r} uses {mode!r} and requires an explicit degradation approval"))
                continue
            used_degradation_ids.add(degradation_id)
            record_entry = degradation_by_id.get(degradation_id)
            if record_entry is None:
                errors.append(_issue("degradation.approval_missing", (collection, index, "reconstruction", "degradation_id"), f"Scene entity {object_id!r} references unknown degradation approval {degradation_id!r}"))
                continue
            approval_index, record = record_entry
            if record.get("approved") is not True:
                errors.append(_issue("degradation.not_approved", ("approvals", "degradations", approval_index, "approved"), f"Degradation {degradation_id!r} for scene entity {object_id!r} is not approved"))
            if object_id not in record.get("entity_ids", []):
                errors.append(_issue("degradation.scope_mismatch", (collection, index, "reconstruction", "degradation_id"), f"Approval {degradation_id!r} does not include scene entity {object_id!r}"))
        elif isinstance(degradation_id, str):
            warnings.append(_issue("degradation.unneeded_reference", (collection, index, "reconstruction", "degradation_id"), f"Native-exact scene entity {object_id!r} carries an unnecessary degradation reference"))

    for approval_id, (approval_index, _) in degradation_by_id.items():
        if approval_id not in used_degradation_ids:
            warnings.append(
                _issue(
                    "degradation.orphaned_approval",
                    ("approvals", "degradations", approval_index),
                    f"Degradation approval {approval_id!r} is not referenced by any degraded scene entity",
                )
            )

    summary = {
        "objects": len(objects),
        "connections": len(connections),
        "degradation_records": len(degradations),
        "degraded_entities": degraded_count,
        "degraded_objects": degraded_object_count,
        "degraded_connections": degraded_connection_count,
        "unsupported_entities": unsupported_count,
        "unsupported_objects": unsupported_object_count,
        "unsupported_connections": unsupported_connection_count,
    }
    return errors, warnings, summary


def _manifest_source(manifest: Any, manifest_path: Path, errors: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not isinstance(manifest, dict):
        errors.append(_issue("source_manifest.type", (), "Source manifest root must be a JSON object", manifest=str(manifest_path)))
        return None
    if manifest.get("kind") != EXPECTED_MANIFEST_KIND:
        errors.append(_issue("source_manifest.kind", ("kind",), f"Expected source manifest kind {EXPECTED_MANIFEST_KIND!r}", actual=manifest.get("kind"), manifest=str(manifest_path)))
    if manifest.get("ok") is not True:
        errors.append(_issue("source_manifest.failed", ("ok",), "Source manifest did not pass preflight", manifest_errors=manifest.get("errors", [])))
    source = manifest.get("source")
    if not isinstance(source, dict):
        errors.append(_issue("source_manifest.source_missing", ("source",), "Source manifest must contain a source object"))
        return None
    for field in ("sha256", "format", "width_px", "height_px"):
        if field not in source:
            errors.append(_issue("source_manifest.field_missing", ("source", field), f"Source manifest is missing source.{field}"))
    if not _is_valid_sha256(source.get("sha256")):
        errors.append(_issue("source_manifest.hash_invalid", ("source", "sha256"), "Source manifest SHA-256 must contain 64 hexadecimal characters"))
    return source


def verify_source(
    plan: dict[str, Any],
    plan_path: Path,
    source_manifest_path: Path | None,
    source_path: Path | None,
    expected_sha256: str | None,
    skip_source_match: bool,
) -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, Any]]]:
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    plan_source = plan.get("source") if isinstance(plan.get("source"), dict) else {}
    planned_hash = plan_source.get("sha256")
    report: dict[str, Any] = {
        "status": "not_checked",
        "method": None,
        "planned_sha256": planned_hash,
        "actual_sha256": None,
    }
    if skip_source_match:
        report.update(status="skipped", method="explicit-skip")
        warnings.append(_issue("source.match_skipped", ("source", "sha256"), "Source identity check was explicitly skipped"))
        return report, errors, warnings
    actual_hash: str | None = None
    manifest_source: dict[str, Any] | None = None
    if source_manifest_path is not None:
        resolved_manifest = source_manifest_path.expanduser().resolve()
        report["method"] = "source-manifest"
        report["source_manifest"] = str(resolved_manifest)
        try:
            manifest = _read_json(resolved_manifest)
        except (OSError, json.JSONDecodeError) as exc:
            errors.append(_issue("source_manifest.read_failed", (), f"Could not read source manifest: {exc}", manifest=str(resolved_manifest)))
        else:
            manifest_source = _manifest_source(manifest, resolved_manifest, errors)
            if manifest_source and _is_valid_sha256(manifest_source.get("sha256")):
                actual_hash = manifest_source["sha256"].lower()
    elif source_path is not None:
        resolved_source = source_path.expanduser().resolve()
        report["method"] = "source-file"
        report["source_path"] = str(resolved_source)
        try:
            actual_hash = _sha256_file(resolved_source)
        except OSError as exc:
            errors.append(_issue("source.read_failed", (), f"Could not hash source image: {exc}", source=str(resolved_source)))
    elif expected_sha256 is not None:
        report["method"] = "expected-sha256"
        actual_hash = expected_sha256.lower()
    elif isinstance(plan_source.get("path"), str):
        candidate = Path(plan_source["path"]).expanduser()
        if not candidate.is_absolute():
            candidate = plan_path.parent / candidate
        candidate = candidate.resolve()
        report["method"] = "plan-source-path"
        report["source_path"] = str(candidate)
        try:
            actual_hash = _sha256_file(candidate)
        except OSError as exc:
            errors.append(_issue("source.read_failed", ("source", "path"), f"Could not hash source image from plan.source.path: {exc}", source=str(candidate)))
    else:
        errors.append(_issue("source.unverified", ("source", "sha256"), "Provide --source-manifest, --source, or --expected-source-sha256; otherwise use --skip-source-match explicitly"))

    report["actual_sha256"] = actual_hash
    if actual_hash is not None and _is_valid_sha256(planned_hash):
        if actual_hash.lower() != planned_hash.lower():
            report["status"] = "fail"
            errors.append(_issue("source.hash_mismatch", ("source", "sha256"), "Scene plan source SHA-256 does not match the locked source", planned=planned_hash.lower(), actual=actual_hash.lower()))
        else:
            report["status"] = "pass"
    elif errors:
        report["status"] = "fail"

    if manifest_source is not None:
        comparisons = (("format", str.upper), ("width_px", lambda item: item), ("height_px", lambda item: item))
        for field, normalize in comparisons:
            planned = plan_source.get(field)
            actual = manifest_source.get(field)
            try:
                matches = normalize(planned) == normalize(actual)
            except (TypeError, AttributeError):
                matches = planned == actual
            if planned is not None and actual is not None and not matches:
                report["status"] = "fail"
                errors.append(_issue("source.metadata_mismatch", ("source", field), f"Scene plan source.{field} differs from source manifest", planned=planned, actual=actual))
    return report, errors, warnings


def _has_panel_lineage(plan_source: dict[str, Any]) -> bool:
    return any(
        field in plan_source
        for field in (
            "parent_sha256",
            "parent_width_px",
            "parent_height_px",
            "panel_manifest_sha256",
            "panel_bbox",
            "panel_label",
            "panel_selection",
        )
    )


def verify_panel_manifest(
    plan: dict[str, Any],
    panel_manifest_path: Path | None,
) -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, Any]]]:
    """Bind mixed-panel lineage claims to panel_crop.py's immutable report."""
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    plan_source = plan.get("source") if isinstance(plan.get("source"), dict) else {}
    lineage_present = _has_panel_lineage(plan_source)
    report: dict[str, Any] = {
        "status": "not_applicable" if not lineage_present else "not_checked",
        "required": lineage_present,
        "path": None,
        "planned_sha256": plan_source.get("panel_manifest_sha256"),
        "actual_sha256": None,
        "comparisons_checked": 0,
    }
    if not lineage_present:
        if panel_manifest_path is not None:
            resolved = panel_manifest_path.expanduser().resolve()
            report.update(status="fail", path=str(resolved))
            errors.append(
                _issue(
                    "panel_manifest.unexpected",
                    (),
                    "--panel-manifest was provided, but the scene plan declares no panel lineage",
                    panel_manifest=str(resolved),
                )
            )
        return report, errors, warnings
    if panel_manifest_path is None:
        report["status"] = "fail"
        errors.append(
            _issue(
                "panel_manifest.required",
                ("source", "panel_manifest_sha256"),
                "A scene plan with parent/panel lineage requires --panel-manifest",
            )
        )
        return report, errors, warnings

    resolved = panel_manifest_path.expanduser().resolve()
    report["path"] = str(resolved)
    try:
        actual_manifest_sha256 = _sha256_file(resolved)
        manifest = _read_json(resolved)
    except (OSError, json.JSONDecodeError) as exc:
        report["status"] = "fail"
        errors.append(_issue("panel_manifest.read_failed", (), f"Could not read panel manifest: {exc}", panel_manifest=str(resolved)))
        return report, errors, warnings
    report["actual_sha256"] = actual_manifest_sha256
    planned_manifest_sha256 = plan_source.get("panel_manifest_sha256")
    if not _is_valid_sha256(planned_manifest_sha256) or planned_manifest_sha256.lower() != actual_manifest_sha256:
        errors.append(
            _issue(
                "panel_manifest.hash_mismatch",
                ("source", "panel_manifest_sha256"),
                "Scene plan panel_manifest_sha256 does not match the supplied manifest file",
                planned=planned_manifest_sha256,
                actual=actual_manifest_sha256,
            )
        )
    if not isinstance(manifest, dict):
        report["status"] = "fail"
        errors.append(_issue("panel_manifest.root_type", (), "Panel manifest root must be a JSON object"))
        return report, errors, warnings
    if manifest.get("kind") != EXPECTED_PANEL_MANIFEST_KIND:
        errors.append(
            _issue(
                "panel_manifest.kind",
                ("kind",),
                f"Panel manifest kind must be {EXPECTED_PANEL_MANIFEST_KIND!r}",
                actual=manifest.get("kind"),
            )
        )
    if manifest.get("tool") != "sci-diagram-pptx/panel_crop.py" or manifest.get("tool_version") != "1.0":
        errors.append(
            _issue(
                "panel_manifest.tool",
                ("tool",),
                "Panel manifest must be produced by sci-diagram-pptx/panel_crop.py version 1.0",
                actual_tool=manifest.get("tool"),
                actual_version=manifest.get("tool_version"),
            )
        )
    if manifest.get("status") != "PASS" or manifest.get("ok") is not True:
        errors.append(
            _issue(
                "panel_manifest.status",
                ("status",),
                "Panel manifest must have status PASS and ok=true",
                actual_status=manifest.get("status"),
                actual_ok=manifest.get("ok"),
            )
        )
    if manifest.get("hard_failures") != []:
        errors.append(
            _issue(
                "panel_manifest.hard_failures",
                ("hard_failures",),
                "Panel manifest contains hard failures",
                hard_failures=manifest.get("hard_failures"),
            )
        )
    if manifest.get("errors") != [] or manifest.get("warnings") != []:
        errors.append(
            _issue(
                "panel_manifest.issues",
                (),
                "A PASS panel manifest must not contain errors or warnings",
                errors=manifest.get("errors"),
                warnings=manifest.get("warnings"),
            )
        )
    manifest_checks = manifest.get("checks")
    if not isinstance(manifest_checks, list) or not manifest_checks or any(
        not isinstance(check, dict) or check.get("status") != "PASS" for check in manifest_checks
    ):
        errors.append(
            _issue(
                "panel_manifest.checks",
                ("checks",),
                "Panel manifest must contain one or more PASS checks and no non-PASS check",
            )
        )

    parent = manifest.get("parent") if isinstance(manifest.get("parent"), dict) else {}
    panel = manifest.get("panel") if isinstance(manifest.get("panel"), dict) else {}
    crop = manifest.get("crop") if isinstance(manifest.get("crop"), dict) else {}
    selection = manifest.get("selection") if isinstance(manifest.get("selection"), dict) else {}
    planned_selection = plan_source.get("panel_selection") if isinstance(plan_source.get("panel_selection"), dict) else {}
    comparisons: list[tuple[str, Any, Any, tuple[Any, ...]]] = [
        ("parent sha256", plan_source.get("parent_sha256"), parent.get("sha256"), ("source", "parent_sha256")),
        ("parent displayed width", plan_source.get("parent_width_px"), parent.get("display_width_px"), ("source", "parent_width_px")),
        ("parent displayed height", plan_source.get("parent_height_px"), parent.get("display_height_px"), ("source", "parent_height_px")),
        ("panel bbox", plan_source.get("panel_bbox"), panel.get("bbox"), ("source", "panel_bbox")),
        ("panel label", plan_source.get("panel_label"), panel.get("label"), ("source", "panel_label")),
        ("crop sha256", plan_source.get("sha256"), crop.get("sha256"), ("source", "sha256")),
        ("crop width", plan_source.get("width_px"), crop.get("width_px"), ("source", "width_px")),
        ("crop height", plan_source.get("height_px"), crop.get("height_px"), ("source", "height_px")),
    ]
    for field in ("selection_source", "selected_by", "selected_at", "evidence"):
        comparisons.append(
            (
                f"selection {field}",
                planned_selection.get(field),
                selection.get(field),
                ("source", "panel_selection", field),
            )
        )
    for label, planned, actual, path in comparisons:
        report["comparisons_checked"] += 1
        if isinstance(planned, str) and isinstance(actual, str) and "sha256" in label:
            matches = planned.lower() == actual.lower()
        else:
            matches = planned == actual
        if not matches:
            errors.append(
                _issue(
                    "panel_manifest.metadata_mismatch",
                    path,
                    f"Scene plan {label} differs from panel manifest",
                    planned=planned,
                    actual=actual,
                )
            )
    report["status"] = "fail" if errors else "pass"
    return report, errors, warnings


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        issue = _issue("cli.invalid_arguments", (), message)
        payload = {"tool": TOOL_NAME, "tool_version": VALIDATOR_VERSION, "validator_version": VALIDATOR_VERSION, "ok": False, "status": "FAIL", "errors": [issue], "hard_failures": [issue], "warnings": [], "checks": []}
        print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        raise SystemExit(2)


def _sha256_arg(value: str) -> str:
    if not _is_valid_sha256(value):
        raise argparse.ArgumentTypeError("must be exactly 64 hexadecimal characters")
    return value.lower()


def _confidence_arg(value: str) -> float:
    try:
        result = float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("must be a number between 0 and 1") from exc
    if not math.isfinite(result) or not 0 <= result <= 1:
        raise argparse.ArgumentTypeError("must be a number between 0 and 1")
    return result


def _write_json(payload: dict[str, Any], output: Path | None, compact: bool) -> None:
    rendered = json.dumps(payload, ensure_ascii=False, indent=None if compact else 2) + "\n"
    if output is None:
        sys.stdout.write(rendered)
        return
    output = output.expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(output.name + ".tmp")
    temporary.write_text(rendered, encoding="utf-8")
    os.replace(temporary, output)


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(description="Validate a sci-diagram-pptx scene plan and emit a JSON report.")
    parser.add_argument("scene_plan", type=Path, help="scene-plan JSON file")
    parser.add_argument("--schema", type=Path, default=Path(__file__).with_name("scene-plan.schema.json"), help="JSON Schema path (default: bundled schema)")
    identity = parser.add_mutually_exclusive_group()
    identity.add_argument("--source-manifest", "--source-preflight", dest="source_manifest", type=Path, help="JSON emitted by source_preflight.py")
    identity.add_argument("--source", type=Path, help="source image to hash directly")
    identity.add_argument("--expected-source-sha256", type=_sha256_arg, help="trusted SHA-256 to compare with plan.source.sha256")
    parser.add_argument("--panel-manifest", type=Path, help="panel_crop.py JSON manifest; required when the plan declares parent/panel lineage")
    parser.add_argument("--skip-source-match", action="store_true", help="explicitly skip source identity verification (warning)")
    parser.add_argument("--min-confidence", type=_confidence_arg, default=0.75, help="fail below this confidence (default: 0.75)")
    confidence_policy = parser.add_mutually_exclusive_group()
    confidence_policy.add_argument("--allow-low-confidence", action="store_true", help="downgrade below-threshold confidence failures to warnings")
    confidence_policy.add_argument("--low-confidence-as-error", action="store_true", help="enforce the default hard-failure policy (compatibility flag)")
    parser.add_argument("--strict", action="store_true", help="make every warning a hard failure")
    parser.add_argument("--output", type=Path, help="write report here instead of stdout")
    parser.add_argument("--compact", action="store_true", help="emit compact JSON")
    return parser


def _empty_report(plan_path: Path, schema_path: Path) -> dict[str, Any]:
    return {
        "tool": TOOL_NAME,
        "tool_version": VALIDATOR_VERSION,
        "validator_version": VALIDATOR_VERSION,
        "kind": "sci-diagram-pptx-scene-plan-validation",
        "checked_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "ok": False,
        "status": "FAIL",
        "scene_plan": str(plan_path),
        "scene_plan_sha256": None,
        "schema": str(schema_path),
        "source_match": {"status": "not_checked", "method": None},
        "panel_manifest_match": {"status": "not_checked", "required": False},
        "summary": {"objects": 0, "connections": 0, "degradation_records": 0, "degraded_entities": 0, "degraded_objects": 0, "degraded_connections": 0, "unsupported_entities": 0, "unsupported_objects": 0, "unsupported_connections": 0},
        "errors": [],
        "hard_failures": [],
        "warnings": [],
        "checks": [],
    }


def _finalize_report(report: dict[str, Any]) -> None:
    report["hard_failures"] = list(report["errors"])
    report["ok"] = not report["hard_failures"]
    report["status"] = "FAIL" if report["hard_failures"] else "WARN" if report["warnings"] else "PASS"


def _semantic_checks(
    plan: dict[str, Any],
    semantic_errors: list[dict[str, Any]],
    semantic_warnings: list[dict[str, Any]],
    minimum_confidence: float,
) -> list[dict[str, Any]]:
    objects = plan.get("objects") if isinstance(plan.get("objects"), list) else []
    connections = plan.get("connections") if isinstance(plan.get("connections"), list) else []
    entities = [item for item in [*objects, *connections] if isinstance(item, dict)]

    def matching(issues: list[dict[str, Any]], prefixes: tuple[str, ...]) -> list[dict[str, Any]]:
        return [item for item in issues if any(item.get("code", "").startswith(prefix) for prefix in prefixes)]

    def check(check_id: str, prefixes: tuple[str, ...], evidence: dict[str, Any], complete: bool = True) -> dict[str, Any]:
        failures = matching(semantic_errors, prefixes)
        cautions = matching(semantic_warnings, prefixes)
        status = "FAIL" if failures or not complete else "WARN" if cautions else "PASS"
        return {
            "id": check_id,
            "status": status,
            "message": f"{check_id} {'passed with evidence' if status == 'PASS' else 'has warnings' if status == 'WARN' else 'failed'}",
            "evidence": {**evidence, "contract_complete": complete, "failure_count": len(failures), "warning_count": len(cautions)},
        }

    object_ids = [item.get("id") for item in objects if isinstance(item, dict) and isinstance(item.get("id"), str)]
    connection_ids = [item.get("id") for item in connections if isinstance(item, dict) and isinstance(item.get("id"), str)]
    normalized_boxes = sum(
        isinstance(item.get(name), dict)
        for item in objects
        if isinstance(item, dict)
        for name in ("source_bbox", "bbox")
    )
    content_contracts = sum(
        isinstance(item.get("content"), dict) and isinstance(item["content"].get("text"), str)
        for item in objects
        if isinstance(item, dict)
    )
    object_style_contracts = sum(
        isinstance(item.get("style"), dict)
        and isinstance(item["style"].get("shape_geometry"), str)
        and bool(item["style"].get("shape_geometry"))
        and all(name not in item["style"] or isinstance(item["style"][name], dict) for name in ("fill", "line", "text"))
        for item in objects
        if isinstance(item, dict)
    )
    z_values = [item.get("z_order") for item in entities if isinstance(item.get("z_order"), int) and not isinstance(item.get("z_order"), bool)]
    confidences = [item.get("confidence") for item in entities if _is_number(item.get("confidence"))]
    endpoint_references = sum(
        isinstance(connection.get(endpoint), dict) and isinstance(connection[endpoint].get("object_id"), str)
        for connection in connections
        if isinstance(connection, dict)
        for endpoint in ("from", "to")
    )
    traced_connections = sum(
        isinstance(connection.get("source_route"), list)
        and len(connection["source_route"]) >= 2
        and isinstance(connection.get("route"), list)
        and len(connection["route"]) >= 2
        for connection in connections
        if isinstance(connection, dict)
    )
    connection_style_contracts = sum(
        isinstance(connection.get("style"), dict)
        and isinstance(connection["style"].get("kind"), str)
        and bool(connection["style"].get("kind"))
        and connection["style"].get("arrow_at") in {"none", "start", "end", "both"}
        and isinstance(connection["style"].get("dash"), str)
        and bool(connection["style"].get("dash"))
        and _is_number(connection["style"].get("line_width_px"))
        and connection["style"]["line_width_px"] > 0
        and isinstance(connection["style"].get("line_color"), str)
        and bool(connection["style"].get("line_color"))
        for connection in connections
        if isinstance(connection, dict)
    )
    approvals = plan.get("approvals") if isinstance(plan.get("approvals"), dict) else {}
    records = approvals.get("degradations") if isinstance(approvals.get("degradations"), list) else []
    degraded_entities = sum(
        isinstance(item.get("reconstruction"), dict) and item["reconstruction"].get("mode") in DEGRADED_MODES
        for item in entities
        if isinstance(item, dict)
    )
    reconstruction_contracts = sum(
        isinstance(item.get("reconstruction"), dict)
        and isinstance(item["reconstruction"].get("mode"), str)
        and isinstance(item["reconstruction"].get("expected_ooxml_kind"), str)
        and bool(item["reconstruction"].get("expected_ooxml_kind"))
        for item in entities
        if isinstance(item, dict)
    )
    complete_approvals = sum(
        isinstance(record, dict)
        and isinstance(record.get("id"), str)
        and bool(record.get("id"))
        and isinstance(record.get("entity_ids"), list)
        and bool(record.get("entity_ids"))
        and isinstance(record.get("reason"), str)
        and bool(record.get("reason"))
        and record.get("approved") is True
        and record.get("approval_source") == "user-explicit"
        and all(isinstance(record.get(field), str) and bool(record.get(field)) for field in ("approved_by", "approved_at", "evidence"))
        for record in records
    )
    panel_fields = sum(
        field in plan.get("source", {})
        for field in (
            "parent_sha256",
            "parent_width_px",
            "parent_height_px",
            "panel_manifest_sha256",
            "panel_bbox",
            "panel_label",
            "panel_selection",
        )
    ) if isinstance(plan.get("source"), dict) else 0
    return [
        check(
            "scene-kind-and-panel-contract",
            ("scene_plan.", "source.panel_"),
            {
                "expected_kind": EXPECTED_PLAN_KIND,
                "actual_kind": plan.get("kind"),
                "panel_metadata_fields_present": panel_fields,
                "panel_lineage_verification": "metadata-consistency-only" if panel_fields else "not-applicable",
            },
        ),
        check(
            "unique-entity-ids",
            ("object.id_", "entity.id_"),
            {"object_ids_checked": len(object_ids), "connection_ids_checked": len(connection_ids), "unique_ids": len(set(object_ids + connection_ids))},
            complete=len(object_ids) == len(objects) and len(connection_ids) == len(connections) and len(set(object_ids + connection_ids)) == len(object_ids) + len(connection_ids),
        ),
        check(
            "normalized-source-and-slide-bboxes",
            ("bbox.",),
            {"objects": len(objects), "bboxes_checked": normalized_boxes, "expected_bboxes": len(objects) * 2},
            complete=normalized_boxes == len(objects) * 2,
        ),
        check(
            "object-content-and-style-contracts",
            ("object.content_", "object.style_"),
            {
                "objects": len(objects),
                "content_contracts_checked": content_contracts,
                "object_style_contracts_checked": object_style_contracts,
            },
            complete=content_contracts == len(objects) and object_style_contracts == len(objects),
        ),
        check(
            "deterministic-z-order",
            ("z_order.",),
            {"entities": len(entities), "z_orders_checked": len(z_values), "unique_z_orders": len(set(z_values))},
            complete=len(z_values) == len(entities) and len(set(z_values)) == len(z_values),
        ),
        check(
            "connection-references",
            ("connection.",),
            {
                "connections": len(connections),
                "endpoint_references_checked": endpoint_references,
                "dual_routes_checked": traced_connections,
                "style_contracts_checked": connection_style_contracts,
            },
            complete=endpoint_references == len(connections) * 2 and traced_connections == len(connections) and connection_style_contracts == len(connections),
        ),
        check(
            "confidence-threshold",
            ("confidence.",),
            {"entities": len(entities), "confidence_values_checked": len(confidences), "minimum": minimum_confidence},
            complete=len(confidences) == len(entities),
        ),
        check(
            "scene-entity-reconstruction-contract",
            ("reconstruction.",),
            {"entities": len(entities), "complete_reconstruction_contracts": reconstruction_contracts},
            complete=reconstruction_contracts == len(entities),
        ),
        check(
            "explicit-degradation-approval",
            ("degradation.", "reconstruction."),
            {"degraded_entities": degraded_entities, "approval_records_checked": len(records), "complete_approval_records": complete_approvals},
            complete=complete_approvals == len(records) and (degraded_entities == 0 or len(records) > 0),
        ),
    ]


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.skip_source_match and any((args.source_manifest, args.source, args.expected_source_sha256)):
        build_parser().error("--skip-source-match cannot be combined with a source identity option")
    plan_path = args.scene_plan.expanduser().resolve()
    schema_path = args.schema.expanduser().resolve()
    report = _empty_report(plan_path, schema_path)
    try:
        try:
            schema = _read_json(schema_path)
        except (OSError, json.JSONDecodeError) as exc:
            report["errors"].append(_issue("schema.read_failed", (), f"Could not read schema: {exc}"))
            report["checks"].append({"id": "schema-load", "status": "FAIL", "message": "Bundled scene-plan schema could not be read", "evidence": {"schema": str(schema_path), "error": str(exc)}})
            _finalize_report(report)
            _write_json(report, args.output, args.compact)
            return 1
        if not isinstance(schema, dict):
            report["errors"].append(_issue("schema.root_type", (), "Schema root must be a JSON object"))
            report["checks"].append({"id": "schema-load", "status": "FAIL", "message": "Scene-plan schema root is not a JSON object", "evidence": {"schema": str(schema_path), "actual_type": type(schema).__name__}})
            _finalize_report(report)
            _write_json(report, args.output, args.compact)
            return 1
        report["checks"].append({"id": "schema-load", "status": "PASS", "message": "Bundled scene-plan schema loaded", "evidence": {"schema": str(schema_path), "schema_id": schema.get("$id"), "draft": schema.get("$schema")}})
        try:
            plan = _read_json(plan_path)
        except (OSError, json.JSONDecodeError) as exc:
            report["errors"].append(_issue("scene_plan.read_failed", (), f"Could not read scene plan: {exc}"))
            report["checks"].append({"id": "scene-plan-load", "status": "FAIL", "message": "Scene plan could not be read as JSON", "evidence": {"scene_plan": str(plan_path), "error": str(exc)}})
            _finalize_report(report)
            _write_json(report, args.output, args.compact)
            return 1
        try:
            report["scene_plan_sha256"] = _sha256_file(plan_path)
        except OSError as exc:
            report["errors"].append(_issue("scene_plan.hash_failed", (), f"Could not hash scene plan: {exc}"))
            report["checks"].append({"id": "scene-plan-load", "status": "FAIL", "message": "Scene plan loaded but could not be identity-hashed", "evidence": {"scene_plan": str(plan_path), "root_type": type(plan).__name__, "hash_error": str(exc)}})
        else:
            report["checks"].append({"id": "scene-plan-load", "status": "PASS", "message": "Scene plan loaded and identity-hashed", "evidence": {"scene_plan": str(plan_path), "root_type": type(plan).__name__, "sha256": report["scene_plan_sha256"]}})
        schema_errors = LightweightSchemaValidator(schema).validate(plan)
        report["errors"].extend(schema_errors)
        report["checks"].append(
            {
                "id": "json-schema-validation",
                "status": "FAIL" if schema_errors else "PASS",
                "message": "Scene plan failed JSON Schema validation" if schema_errors else "Scene plan passed JSON Schema validation",
                "evidence": {"schema_error_count": len(schema_errors), "schema": str(schema_path)},
            }
        )
        if isinstance(plan, dict):
            semantic_errors, semantic_warnings, summary = validate_semantics(plan, args.min_confidence, not args.allow_low_confidence)
            report["errors"].extend(semantic_errors)
            report["warnings"].extend(semantic_warnings)
            report["summary"] = summary
            report["checks"].extend(_semantic_checks(plan, semantic_errors, semantic_warnings, args.min_confidence))
            source_report, source_errors, source_warnings = verify_source(
                plan,
                plan_path,
                args.source_manifest,
                args.source,
                args.expected_source_sha256,
                args.skip_source_match,
            )
            report["source_match"] = source_report
            report["errors"].extend(source_errors)
            report["warnings"].extend(source_warnings)
            source_check_status = "PASS" if source_report.get("status") == "pass" and not source_errors else "WARN" if source_report.get("status") == "skipped" and not source_errors else "FAIL"
            report["checks"].append(
                {
                    "id": "source-identity-match",
                    "status": source_check_status,
                    "message": "Scene plan source identity matches the locked source" if source_check_status == "PASS" else "Source identity verification was skipped" if source_check_status == "WARN" else "Scene plan source identity does not match the locked source",
                    "evidence": {
                        "method": source_report.get("method"),
                        "planned_sha256": source_report.get("planned_sha256"),
                        "actual_sha256": source_report.get("actual_sha256"),
                        "source_error_count": len(source_errors),
                    },
                }
            )
            panel_report, panel_errors, panel_warnings = verify_panel_manifest(plan, args.panel_manifest)
            report["panel_manifest_match"] = panel_report
            report["errors"].extend(panel_errors)
            report["warnings"].extend(panel_warnings)
            panel_status = panel_report.get("status")
            panel_check_status = "PASS" if panel_status in {"pass", "not_applicable"} and not panel_errors else "FAIL"
            report["checks"].append(
                {
                    "id": "panel-lineage-manifest",
                    "status": panel_check_status,
                    "message": (
                        "Panel lineage manifest matches the scene plan"
                        if panel_status == "pass" and panel_check_status == "PASS"
                        else "Panel lineage manifest is not applicable to this standalone source"
                        if panel_status == "not_applicable" and panel_check_status == "PASS"
                        else "Panel lineage manifest is missing or inconsistent with the scene plan"
                    ),
                    "evidence": {
                        "required": panel_report.get("required"),
                        "path": panel_report.get("path"),
                        "planned_sha256": panel_report.get("planned_sha256"),
                        "actual_sha256": panel_report.get("actual_sha256"),
                        "comparisons_checked": panel_report.get("comparisons_checked"),
                        "failure_count": len(panel_errors),
                    },
                }
            )
        else:
            report["checks"].append({"id": "semantic-validation", "status": "FAIL", "message": "Semantic validation requires a JSON object root", "evidence": {"reason": "scene plan root is not an object"}})
        if args.strict and report["warnings"]:
            report["errors"].append(_issue("strict.warnings_present", (), "Warnings are errors in --strict mode", warning_count=len(report["warnings"])))
        _finalize_report(report)
        _write_json(report, args.output, args.compact)
        return 0 if report["ok"] else 1
    except OSError as exc:
        issue = _issue("io.write_failed", (), str(exc))
        payload = {"tool": TOOL_NAME, "tool_version": VALIDATOR_VERSION, "validator_version": VALIDATOR_VERSION, "ok": False, "status": "FAIL", "errors": [issue], "hard_failures": [issue], "warnings": [], "checks": []}
        print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
