---
name: ppt-animation-coach
description: Create reusable PowerPoint animation techniques, including a five-image carousel and a three-line text progression with editable objects and Morph transitions. Use when a user wants one of these learned animation patterns; do not use for ordinary static slide design or unrelated animations.
---

# PPT Animation Coach

This skill currently contains two techniques:

- Image carousel: create a playable Microsoft PowerPoint carousel from exactly five user-supplied images.
- Three-line text progression: create a three-slide text reveal and repositioning sequence from three user-supplied text lines.

Keep all objects editable and explain the construction in slide notes.

## Workflow

1. Require exactly five image paths. If fewer are supplied, ask for the missing images. If more are supplied, use the first five only after stating that choice.
2. Run [scripts/build-image-carousel.ps1](scripts/build-image-carousel.ps1) or reproduce its construction in a connected PowerPoint session.
3. Create five rounded picture frames: one large center card, two medium cards, and two small outer cards. Add subtle shadows, reflection, and no outline.
4. Duplicate the entire first slide four times. Do not copy objects from the canvas into an existing slide, because new object IDs can break Morph matching.
5. Across slides, rotate the same five picture frames one position clockwise. Keep the visual stack explicit: center in front, medium cards behind it, and the far-right small card at the bottom when it overlaps.
6. Use PowerPoint Morph (`ppEffectMorphByObject = 3954`) for slides 2 onward, with automatic advance and a default two-second duration. Keep click advance enabled unless unattended playback is requested.
7. Put the construction steps, object-correspondence warning, timing, and WPS caveat in speaker notes. Deliver the `.pptx` and inspect at least the first and last slide in PowerPoint.

## Three-line text progression

1. Require exactly three text lines and preserve them as three text boxes on every slide.
2. Start with the first line visible and the next two lines hidden through text-fill transparency.
3. Duplicate the entire previous slide for each next state. Do not recreate or paste the text boxes into a different slide.
4. On each new slide, move the existing text boxes as a group and reveal the next line by changing only its text-fill transparency.
5. Apply PowerPoint Morph (`ppEffectMorphByObject = 3954`) from slide 2 onward, with automatic advance and a default duration near two seconds.
6. Keep the text-box object names and ordering stable so Morph can match the same text objects across slides.
7. Put the three-line construction logic, transparency rule, timing, and WPS caveat in speaker notes. Deliver the `.pptx` and preview it before handoff.

## Source handling

Use the user's supplied local tutorial video as the technique source. Do not copy the creator's original assets; use the user's own five images or other licensed images. Keep the source record in [references/image-carousel.md](references/image-carousel.md).
