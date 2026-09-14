#!/usr/bin/env python3
"""Create middle-focus and foreground-defocus PNGs from one clean RGBA image."""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
from PIL import Image, ImageFilter


def premultiplied_blur(image: Image.Image, radius: float) -> Image.Image:
    rgba = np.asarray(image.convert("RGBA"), dtype=np.float32) / 255.0
    alpha = rgba[..., 3].copy()
    alpha[alpha < 0.012] = 0
    premultiplied = rgba[..., :3] * alpha[..., None]

    channels = []
    for index in range(3):
        channel = Image.fromarray(np.clip(premultiplied[..., index] * 255, 0, 255).astype(np.uint8), "L")
        channels.append(np.asarray(channel.filter(ImageFilter.GaussianBlur(radius)), dtype=np.float32) / 255.0)

    alpha_image = Image.fromarray(np.clip(alpha * 255, 0, 255).astype(np.uint8), "L")
    blurred_alpha = np.asarray(alpha_image.filter(ImageFilter.GaussianBlur(radius)), dtype=np.float32) / 255.0
    blurred_alpha[blurred_alpha < 0.012] = 0

    blurred_premultiplied = np.stack(channels, axis=2)
    rgb = np.clip(blurred_premultiplied / np.maximum(blurred_alpha[..., None], 1e-5), 0, 1)
    result = Image.fromarray(
        np.clip(np.dstack([rgb, blurred_alpha]) * 255, 0, 255).astype(np.uint8),
        "RGBA",
    )
    bbox = result.getchannel("A").getbbox()
    return result.crop(bbox) if bbox else result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path, help="Clean transparent PNG")
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--mid-radius", type=float)
    parser.add_argument("--front-radius", type=float)
    args = parser.parse_args()

    image = Image.open(args.input).convert("RGBA")
    longest = max(image.size)
    mid_radius = args.mid_radius if args.mid_radius is not None else longest * 0.004
    front_radius = args.front_radius if args.front_radius is not None else longest * 0.016

    args.output_dir.mkdir(parents=True, exist_ok=True)
    stem = args.input.stem
    sharp_path = args.output_dir / f"{stem}-sharp.png"
    mid_path = args.output_dir / f"{stem}-mid-soft.png"
    front_path = args.output_dir / f"{stem}-front-blur.png"
    image.save(sharp_path)
    premultiplied_blur(image, mid_radius).save(mid_path)
    premultiplied_blur(image, front_radius).save(front_path)
    print(sharp_path)
    print(mid_path)
    print(front_path)


if __name__ == "__main__":
    main()

