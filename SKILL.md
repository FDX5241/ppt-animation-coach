---
name: ppt-animation-coach
description: Create a five-image PowerPoint carousel with editable picture frames and Morph transitions. Use when a user wants images to rotate through a layered card layout; do not use for ordinary static slide design or unrelated animations.
---

# PPT Animation Coach

Create a playable Microsoft PowerPoint image carousel from exactly five user-supplied images. Keep the picture frames editable and explain the construction in slide notes.

## Workflow

1. Require exactly five image paths. If fewer are supplied, ask for the missing images. If more are supplied, use the first five only after stating that choice.
2. Run [scripts/build-image-carousel.ps1](scripts/build-image-carousel.ps1) or reproduce its construction in a connected PowerPoint session.
3. Create five rounded picture frames: one large center card, two medium cards, and two small outer cards. Add subtle shadows, reflection, and no outline.
4. Duplicate the entire first slide four times. Do not copy objects from the canvas into an existing slide, because new object IDs can break Morph matching.
5. Across slides, rotate the same five picture frames one position clockwise. Keep the visual stack explicit: center in front, medium cards behind it, and the far-right small card at the bottom when it overlaps.
6. Use PowerPoint Morph (`ppEffectMorphByObject = 3954`) for slides 2 onward, with automatic advance and a default two-second duration. Keep click advance enabled unless unattended playback is requested.
7. Put the construction steps, object-correspondence warning, timing, and WPS caveat in speaker notes. Deliver the `.pptx` and inspect at least the first and last slide in PowerPoint.

## Source handling

Use the user's supplied local tutorial video as the technique source. Do not copy the creator's original assets; use the user's own five images or other licensed images. Keep the source record in [references/image-carousel.md](references/image-carousel.md).
