# Flowing Wave Animation

## Technique summary

This technique creates a continuous water or fluid motion effect with editable
PowerPoint shapes. The demo uses repeated oval wave crests plus a filled base,
then shifts the wave group by a consistent amount across short timed slides.

## When to use it

Use it for water, fluid, atmosphere, technology backgrounds, progress states, or
section transitions. Keep the wave low on the canvas when the slide still needs
to carry a title or an explanation.

## Native construction

1. Insert one oval and duplicate it into a repeated row.
2. Place a rectangle below the ovals so the lower edge becomes a filled wave band.
3. Group the band and make a second band with a different color and vertical
   offset for depth.
4. Duplicate the slide several times and shift both groups horizontally by a
   consistent distance.
5. Turn off advance-on-click, set automatic advance to a short interval, and
   enable looping when the wave should run continuously.

The included demo uses six frames and 0.35 seconds per frame. The horizontal
offset wraps within one crest width so the frame edges do not reveal a hard
vertical seam.

## User-adjustable parameters

- Frame interval: controls perceived speed.
- Shift distance: controls movement per frame and visual smoothness.
- Crest width and height: controls the wave rhythm.
- Layer color, transparency, and vertical offset: controls depth.
- Looping: controls whether the animation repeats during the slideshow.

## Compatibility and limitations

The demo uses editable shapes and timed slide advance, so it should be more
portable than a video-based effect. PowerPoint remains the authoring target;
test the final deck in WPS because automatic advance and looping details may
vary. This implementation is a frame-based animation across slides, not a
single-slide motion-path loop.

## Example

- [wave-animation-demo.pptx](../examples/wave-animation-demo.pptx)
- Builder: `scripts/build-wave-demo.ps1`

## Source record

- Title: PPT制作 | 海浪动画
- Creator: 星瀚文化课件定制
- URL: https://www.bilibili.com/video/BV1kX4y1f7EX/
- Reviewed: 2026-09-13
