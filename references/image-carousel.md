# Image Carousel with Morph

## Visual result

Five rounded picture cards form a layered carousel: one large card in the center, two medium cards beside it, and two smaller cards at the outer edges. Each slide keeps the same five frame objects and rotates the images one position clockwise. Morph animates the cards and their picture fills between slides.

## Construction rules

1. Create the five rounded picture frames and shadows on the first slide.
2. Apply picture fill, reflection, and no outline to each frame.
3. Duplicate the entire slide four times. Do not copy all objects from the canvas into another existing slide, because PowerPoint can assign new object IDs and Morph will lose correspondence.
4. Change only the picture fill assigned to each named frame on each duplicate. Use `new slot = old slot + 1`, wrapping after the rightmost slot.
5. Put the center card in front. Put the far-right small card at the bottom when it overlaps the next card. Keep shadows behind their matching cards.

## Default parameters

- Five images and five slides
- Morph on slides 2 to 5
- Two-second Morph duration
- Optional automatic advance after two seconds
- Rounded cards with reflection when PowerPoint exposes the control

## Example

- [image-carousel-demo.pptx](../examples/image-carousel-demo.pptx)
- [image-carousel-preview.png](../examples/image-carousel-preview.png)

## Failure modes

- Pasting individual objects into an existing slide breaks object matching.
- Reordering or renaming frames inconsistently can produce a jump or fade.
- Sending the center card behind side cards makes the focal image appear clipped.
- WPS may interpret Morph and automatic advance differently from PowerPoint.

## Source record

- Title: `[第三期] PPT教学——图片轮播效果`
- Creator: user-supplied local video; creator attribution was not available in the provided file metadata
- URL: local file supplied by the user, not publicly published
- Reviewed: 2026-09-14
