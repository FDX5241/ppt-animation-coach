---
name: ppt-animation-coach
description: Create or adapt playable PowerPoint animations from text, PPTX files, slide screenshots, reference images, or accessible video examples, while preserving the original slide and teaching the technique. Use for PowerPoint animation implementation or reconstruction; do not use for ordinary static slide design.
---

# PPT Animation Coach

Create the animation first, then provide the teaching material as a supporting
deliverable. Author for Microsoft PowerPoint and make a best-effort check for WPS
playback compatibility.

## Workflow

1. Inspect the user's input and identify the intended visual behavior, timing,
   interaction, and presentation context.
2. If the user did not specify whether to edit an existing deck or create an
   example, ask which outcome they want. If no deck is supplied, ask the intended
   use before designing the example slide.
3. For a supplied deck, preserve the original slide and create the animated
   version as a duplicate unless the user explicitly asks otherwise.
4. Match the request to the technique index in
   [references/technique-index.md](references/technique-index.md). Read only the
   technique references needed for the current request.
   The completed flowing-wave technique is documented in
   [references/wave-animation.md](references/wave-animation.md).
5. Prefer editable PowerPoint shapes, text, images, transitions, and animations.
   Use GIF or video only when a faithful native implementation is impractical.
6. When reconstructing a reference video, first decide whether PowerPoint can
   reproduce the effect reliably. If not, explain the limitation and offer an
   original approximation plus practical alternatives.
7. Deliver a playable `.pptx`. Explain the result in chat and add concise
   reproduction steps, parameters, and a practice suggestion to slide notes.
8. Verify that the animation plays as intended in PowerPoint. Check WPS playback
   where the available environment allows it, and disclose untested or known
   compatibility risks.

## Source handling

Use public tutorials to understand techniques, not to copy protected videos,
transcripts, paid materials, or creator assets. Record technique sources only in
the GitHub reference document, including creator, URL, and review date.

New user-submitted techniques remain drafts until the repository maintainer
reviews them.
