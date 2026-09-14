---
name: ppt-animation-coach
description: Create reusable Microsoft PowerPoint animations with editable objects and Morph transitions. Use for the five-image carousel, three-line text progression, or layered parallax scene; do not use for ordinary static slide design.
---

# PPT Animation Coach

Route the request to one implemented technique:

- Five-image carousel: read [references/image-carousel.md](references/image-carousel.md).
- Three-line text progression: read [references/text-morph-3-line.md](references/text-morph-3-line.md).
- Layered parallax Morph: read [references/parallax-morph.md](references/parallax-morph.md).
- Learning a new animation from a tutorial video: read [references/learn-animation-from-video.md](references/learn-animation-from-video.md).

Preserve editable PowerPoint objects and use full-slide duplication whenever Morph depends on object correspondence. Do not paste a canvas selection into an existing slide.

## Asset approval workflow

Before building any technique that depends on user-supplied or generated inputs:

1. State which assets the selected technique requires and which assets are optional.
2. Inventory the supplied images, text, and other content. Identify missing, unsuitable, low-resolution, non-transparent, or inconsistent items.
3. Tell the user exactly what should be generated or sourced for every missing item. Offer to find or generate those assets when authorized.
4. Prepare the candidate materials and show them to the user. Include text order, planned crops, and derived visual variants that materially affect the animation, such as the middle-focus and foreground-blur versions used by the parallax technique.
5. Wait for the user to approve the assets before constructing the PowerPoint. If the user rejects an asset, revise only the rejected asset and show it again.

This approval workflow applies to every current and future animation technique in this skill, including image carousels and text-only animations. Skip the approval pause only when the user explicitly asks to proceed without previewing materials.

## Developing a new technique from video

When the user supplies a new tutorial video, do not jump directly to building a deck. Follow this order:

1. Extract one representative frame per second for the relevant section.
2. Describe the complete effect and construction logic in language, including every slide state, object count, object operation, layer, timing, and correspondence requirement.
3. Identify, source, generate, and preview all required assets. Wait for user approval.
4. Build the animation, preview it, compare it with the tutorial, and iterate.

Only add a newly learned technique to the formal library after the user accepts a working demonstration.

## Shared delivery rules

- Keep object names stable across duplicated slides.
- Apply PowerPoint Morph (`ppEffectMorphByObject = 3954`) only after duplicating the complete source slide.
- Put the construction logic, timing, source record, and known WPS limitations in speaker notes.
- Preview the finished animation in Microsoft PowerPoint before handoff.
- Deliver the `.pptx`; add an MP4 preview when the user wants an easy review format.
