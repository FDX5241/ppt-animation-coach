# PPT Animation Coach

PPT Animation Coach is a Codex skill for creating playable PowerPoint animations
from a text description, an existing `.pptx`, a slide screenshot, a reference
image, or an accessible video link. It prioritizes editable PowerPoint objects,
preserves the original slide before creating an animated version, and explains
the technique in chat and in slide notes.

The first release contains one focused technique: an editable five-image carousel built with duplicated slides and Morph transitions.

PowerPoint is the authoring target. WPS playback compatibility is a best-effort
goal. When a native editable implementation is impractical, the skill should
explain the limitation and offer a GIF or video fallback.

## Project status

The skill accepts five image paths and can build a playable `.pptx` locally on a
Windows computer with Microsoft PowerPoint installed. See
[`references/image-carousel.md`](references/image-carousel.md) and
[`scripts/build-image-carousel.ps1`](scripts/build-image-carousel.ps1).

## Sources and contributions

Each technique reference records its tutorial sources, author, URL, and review
date. User-submitted techniques enter a draft state and are added to the formal
library only after maintainer review.
