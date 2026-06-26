---
name: ai-comic-drama-production
description: Create, optimize, and review AI comic-drama production assets. Use when Codex needs to turn scripts into AI漫剧 video prompts, 15-second shot segments, nine-grid storyboards, character/space continuity locks, dialogue-preserving prompt templates, or 9.6+ quality reviews.
---

# AI Comic Drama Production

Use this skill as a production workflow for AI漫剧, AI short-drama, cinematic storyboard, and AI video prompt generation.

## Core Rules

Preserve the script first. Do not rewrite, delete, reorder, or add user-provided dialogue. If a line is long, place reaction shots around it while keeping the dialogue text and order unchanged.

Treat every video segment as independent. Rebuild character identity, appearance, position, scene anchors, start pose, end pose, lighting, and continuity constraints inside each segment.

Use screen-relative coordinates when the user provides references, storyboards, grids, or says "以画面为参照": screen left/right/up/down and foreground/midground/background. Avoid ambiguous character-left or character-right language unless also anchored to screen direction.

Default each 15-second story beat to 5-8 cinematic shots. Include focal length, camera height/angle, shot size, movement, action start/process/end, eyeline, micro-expression, lighting contact, and transition logic.

Limit focus characters to 3-4 per independent segment. Treat extras as blurred silhouettes, shoulders, hands, reflections, backlit outlines, or background shapes.

## Workflow

1. Identify the output type:
   - Complete AI video prompt, shot split, final prompt, or 9.6+ version: read `references/prompt-contract.md`.
   - Nine-grid storyboard, image prompt grid, or standing map: read `references/prompt-contract.md`.
   - Analysis, scoring, diagnosis, or improvement advice: read `references/qa-scorecard.md`.
2. Extract immutable inputs: dialogue, character list, plot beat, style, duration, model constraints, reference images, safety requirements, and platform restrictions.
3. Build a spatial map before writing: who is screen left/right/foreground/background, who faces whom, distances, props, doors, screens, cracks, vehicles, weapons, or platforms.
4. Split the beat into cinematic shots. Preserve action continuity with action matching, eyeline matching, foreground occlusion, reaction shots, and axis control.
5. Add performance notes to every dialogue line: breath, pause, stress, sentence landing, rhythm, gaze, micro-expression, and emotional layer.
6. Add technical locks: interaction lock, continuity lock, keyframes, negative constraints, style, sound, lighting, and rhythm.
7. Audit before final output: dialogue fidelity, continuity, physical motion, axis, character count, AI failure risks, moderation-safe wording, and completeness.

## Cinematic Defaults

Prefer ASC-style visual language: Steadicam, dolly track, Technocrane, crash zoom, match cut, macro insert, low-angle pressure shot, over-the-shoulder shot, reaction shot, negative space, foreground occlusion, low-key lighting, rim light, negative fill, ambient occlusion, volumetric light, and cool/warm contrast.

Default style if the user gives none: 3D anime quality, 3D game CG, UE5 render, PBR materials, cinematic virtual photography, Hollywood-level visual effects, ASC cinematography standard, high-frequency short-drama rhythm, no background music, no external subtitles.

If the user asks for live-action realism, switch to: live-action realism, Hollywood cinematic VFX, ASC cinematography standard, premium practical-lighting texture.

For darker, violent, horror, fantasy, or combat material, preserve tension with safer visual wording. Prefer phrases such as "深渊巨影", "毁灭性压迫", "暗红能量痕迹", "倒伏身影", or "封锁危险区" when direct wording may trigger moderation.

## Output Discipline

When the user asks for "最终版", "最高评分版", "9.6+", or "完整", output the complete usable artifact. Do not summarize modules, write "same as above", or omit required sections.

Do not include process explanation unless the user asks for analysis. Deliver production-ready prompt text directly.
