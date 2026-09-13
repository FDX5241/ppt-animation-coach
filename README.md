# PPT Animation Coach

PPT Animation Coach is a Codex skill for creating playable PowerPoint animations
from a text description, an existing `.pptx`, a slide screenshot, a reference
image, or an accessible video link. It prioritizes editable PowerPoint objects,
preserves the original slide before creating an animated version, and explains
the technique in chat and in slide notes.

The first release is planned around ten techniques:

1. Flowing wave
2. Morph transition
3. Mask reveal
4. Liquid fill
5. Motion path
6. Flip card
7. Rolling numbers
8. Countdown
9. Progress bar
10. Image carousel

PowerPoint is the authoring target. WPS playback compatibility is a best-effort
goal. When a native editable implementation is impractical, the skill should
explain the limitation and offer a GIF or video fallback.

## Project status

The repository is in its initial design phase. The skill entry point and
technique routing are being established before animation-generation scripts and
example presentations are added.

## Sources and contributions

Each technique reference records its tutorial sources, author, URL, and review
date. User-submitted techniques enter a draft state and are added to the formal
library only after maintainer review.
