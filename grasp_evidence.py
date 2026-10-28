"""Read numeric grasp evidence and project it without running workspace code."""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
from PIL import Image

# Shared sequential palette for raw and completed depth, including the legend.
DEPTH_COLORS = np.array([[68, 1, 84], [59, 82, 139], [33, 145, 140],
                         [94, 201, 98], [253, 231, 37]], dtype=float)


def depth_colors(values):
    """Map normalized values in [0, 1] to the same RGB palette."""
    values = np.clip(values, 0, 1)
    stops = np.linspace(0, 1, len(DEPTH_COLORS))
    return np.stack([np.interp(values, stops, DEPTH_COLORS[:, i])
                     for i in range(3)], axis=-1).astype(np.uint8)


def depth_pair(before: Path, after: Path, crop: list[int], limits: list[float]):
    """Return two atomic depth images using identical pixel ROI and meter limits."""
    arrays = [np.load(path, allow_pickle=False) for path in (before, after)]
    if arrays[0].ndim != 2 or arrays[0].shape != arrays[1].shape:
        raise ValueError("depth pair requires matching two-dimensional arrays")
    left, top, right, bottom = crop
    height, width = arrays[0].shape
    if not (0 <= left < right <= width and 0 <= top < bottom <= height):
        raise ValueError("depth crop exceeds source dimensions")
    low, high = limits
    if not np.isfinite([low, high]).all() or not 0 < low < high:
        raise ValueError("depth pair requires increasing positive meter limits")
    pictures = []
    for array in arrays:
        roi = array[top:bottom, left:right]
        valid = np.isfinite(roi) & (roi > 0)
        normalized = np.zeros(roi.shape, dtype=float)
        normalized[valid] = (roi[valid] - low) / (high - low)
        rgb = depth_colors(normalized)
        rgb[~valid] = [190, 190, 190]
        pictures.append(Image.fromarray(rgb))
    return pictures


def gripper_projection(execution: Path, calibration: Path, tip_z_m=.1034, finger_length_m=.04,
                       finger_width_m=.02, finger_thickness_m=.01):
    """Project a selected pose and predicted jaw opening into source RGB pixels.

    The wireframe follows the recorded camera transform. Finger length and
    tip offset describe the display geometry, not a collision or safety test.
    """
    record = json.loads(execution.read_text(encoding="utf-8"))
    camera = json.loads(calibration.read_text(encoding="utf-8"))
    return project_gripper(record, camera, tip_z_m, finger_length_m,
                           finger_width_m, finger_thickness_m)


def candidate_projections(response: Path, calibration: Path, count=None):
    """Project the highest-score saved candidates with the selected-pose geometry."""
    record = json.loads(response.read_text(encoding="utf-8"))
    camera = json.loads(calibration.read_text(encoding="utf-8"))
    grasps = record["grasps"]
    candidates = []
    if count is not None and (type(count) is not int or not 1 <= count <= len(grasps.get("scores", {}).get(next(iter(grasps["scores"]), ""), []))):
        raise ValueError("candidate display count is outside the recorded candidates")
    for group, poses in grasps["pred_grasps_cam"].items():
        scores, openings = grasps["scores"][group], grasps["gripper_openings"][group]
        if len(poses) != len(scores) or len(poses) != len(openings):
            raise ValueError("candidate poses, scores and openings must have matching lengths")
        for index, (pose, score, opening) in enumerate(zip(poses, scores, openings)):
            if type(score) not in {int, float} or not np.isfinite(score):
                raise ValueError("candidate scores must be finite numbers")
            candidates.append((score, group, index, pose, opening))
    if len(candidates) != grasps["total_grasps"] or not candidates:
        raise ValueError("candidate arrays must match the recorded nonzero total")
    projections = []
    ranked = sorted(candidates, key=lambda c: -c[0])
    for score, group, index, pose, opening in ranked if count is None else ranked[:count]:
        projection = project_gripper({"T_camera_cgn_gripper": pose,
                                      "gripper_opening_m": opening}, camera)
        projection.update(score=score, group=group, index=index)
        projections.append(projection)
    return {"total": len(candidates), "projections": projections}


def gated_candidate_projections(execution: Path, calibration: Path):
    """Project every saved candidate and retain its accepted/rejected status."""
    record = json.loads(execution.read_text(encoding="utf-8"))
    camera = json.loads(calibration.read_text(encoding="utf-8"))
    accepted = record.get("candidate_acceptances", [])
    rejected = record.get("candidate_rejections", [])
    candidates = [(item, "accepted") for item in accepted] + [(item, "rejected") for item in rejected]
    expected = record.get("candidates_evaluated")
    if type(expected) is not int or expected <= 0 or len(candidates) != expected:
        raise ValueError("candidate gate record must contain every evaluated candidate")
    projections = []
    for item, status in candidates:
        pose = item.get("T_camera_cgn_gripper")
        opening = item.get("gripper_opening_m")
        if pose is None or opening is None:
            raise ValueError("candidate gate record is missing pose or opening")
        projection = project_gripper({"T_camera_cgn_gripper": pose,
                                      "gripper_opening_m": opening}, camera)
        projection.update(status=status, index=item.get("candidate_index"),
                          score=item.get("candidate_score"))
        projections.append(projection)
    return {"total": expected, "accepted": len(accepted),
            "rejected": len(rejected), "projections": projections}


def project_gripper(record, camera, tip_z_m=.1034, finger_length_m=.04,
                    finger_width_m=.02, finger_thickness_m=.01):
    """Project one recorded pose with common display geometry for every layout."""
    transform = np.asarray(record["T_camera_cgn_gripper"], dtype=float)
    intrinsic = np.asarray(camera["color_k"], dtype=float)
    opening = float(record["gripper_opening_m"])
    if (transform.shape != (4, 4) or intrinsic.shape != (3, 3)
            or not np.isfinite(transform).all() or not np.isfinite(intrinsic).all()
            or not np.isfinite(opening) or not 0 <= opening <= .2
            or not 0 < finger_length_m < tip_z_m
            or not 0 < finger_width_m < .1 or not 0 < finger_thickness_m < .1):
        raise ValueError("invalid recorded gripper pose or display geometry")
    base_z = tip_z_m - finger_length_m
    local = np.array([[-opening / 2, 0, base_z], [-opening / 2, 0, tip_z_m],
                      [opening / 2, 0, base_z], [opening / 2, 0, tip_z_m],
                      [0, 0, tip_z_m]])
    edges = [(0, 1), (0, 2), (2, 3)]
    # Draw each finger as a wireframe volume so an oblique camera view does
    # not reduce the grasp display to three nearly coincident center lines.
    for low_x, high_x in [(-opening / 2 - finger_thickness_m, -opening / 2),
                          (opening / 2, opening / 2 + finger_thickness_m)]:
        start = len(local)
        corners = [[x, y, z] for x in (low_x, high_x)
                   for y in (-finger_width_m / 2, finger_width_m / 2)
                   for z in (base_z, tip_z_m)]
        local = np.concatenate([local, np.array(corners)])
        for a in range(8):
            for axis_bit in (1, 2, 4):
                b = a ^ axis_bit
                if a < b:
                    edges.append((start + a, start + b))
    points = local @ transform[:3, :3].T + transform[:3, 3]
    if (points[:, 2] <= 0).any():
        raise ValueError("gripper projection requires positive camera depth")
    pixels = points @ intrinsic.T
    pixels = pixels[:, :2] / pixels[:, 2:3]
    segments = [(pixels[a].tolist(), pixels[b].tolist()) for a, b in edges]
    return {"segments": segments, "center_uv": pixels[4].tolist(), "opening_mm": opening * 1000}
