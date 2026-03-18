"""Image compression utilities for multimodal support."""

import base64
import io
import logging

logger = logging.getLogger(__name__)


def compress_image(
    image_bytes: bytes,
    max_size_bytes: int,
    max_dimension: int = 1024,
    quality: int = 75,
) -> bytes:
    """Compress image to fit within size limit.

    Args:
        image_bytes: Original image bytes
        max_size_bytes: Maximum allowed size in bytes
        max_dimension: Maximum width/height in pixels
        quality: Initial JPEG quality (1-100)

    Returns:
        Compressed image bytes
    """
    try:
        from PIL import Image
    except ImportError:
        logger.warning("Pillow not installed, returning original image")
        return image_bytes

    try:
        # Open image
        img = Image.open(io.BytesIO(image_bytes))
        original_format = img.format
        original_size = len(image_bytes)
        logger.info(f"[IMAGE COMPRESS] Original: {original_size} bytes, format: {original_format}, size: {img.size}")

        # Convert to RGB if necessary (for JPEG compression)
        if img.mode in ("RGBA", "P", "LA"):
            # Create white background for transparent images
            background = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "P":
                img = img.convert("RGBA")
            if img.mode in ("RGBA", "LA"):
                background.paste(img, mask=img.split()[-1])  # Use alpha channel as mask
                img = background
            else:
                img = img.convert("RGB")

        # Step 1: Resize if needed
        width, height = img.size
        if width > max_dimension or height > max_dimension:
            # Calculate new size maintaining aspect ratio
            if width > height:
                new_width = max_dimension
                new_height = int(height * (max_dimension / width))
            else:
                new_height = max_dimension
                new_width = int(width * (max_dimension / height))
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            logger.info(f"[IMAGE COMPRESS] Resized to: {img.size}")

        # Step 2: Try initial quality
        current_quality = quality
        compressed_bytes = _save_as_jpeg(img, current_quality)

        # Step 3: Reduce quality if still too large
        while len(compressed_bytes) > max_size_bytes and current_quality > 20:
            current_quality -= 10
            compressed_bytes = _save_as_jpeg(img, current_quality)
            logger.info(f"[IMAGE COMPRESS] Quality reduced to {current_quality}, size: {len(compressed_bytes)}")

        # Step 4: Further resize if still too large
        resize_factor = 0.8
        while len(compressed_bytes) > max_size_bytes and min(img.size) > 100:
            new_width = int(img.size[0] * resize_factor)
            new_height = int(img.size[1] * resize_factor)
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            compressed_bytes = _save_as_jpeg(img, current_quality)
            logger.info(f"[IMAGE COMPRESS] Further resized to: {img.size}, size: {len(compressed_bytes)}")
            resize_factor *= 0.9  # Gradually reduce resize factor

        final_size = len(compressed_bytes)
        compression_ratio = (1 - final_size / original_size) * 100
        logger.info(f"[IMAGE COMPRESS] Final: {final_size} bytes, compressed {compression_ratio:.1f}%")

        return compressed_bytes

    except Exception as e:
        logger.error(f"[IMAGE COMPRESS] Error compressing image: {e}")
        return image_bytes


def _save_as_jpeg(img, quality: int) -> bytes:
    """Save image as JPEG with given quality."""
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=quality, optimize=True)
    return buffer.getvalue()


def compress_base64_image(
    base64_content: str,
    max_size_bytes: int,
    max_dimension: int = 1024,
    quality: int = 75,
) -> str:
    """Compress a base64-encoded image.

    Args:
        base64_content: Base64-encoded image content
        max_size_bytes: Maximum allowed size in bytes
        max_dimension: Maximum width/height in pixels
        quality: Initial JPEG quality (1-100)

    Returns:
        Compressed base64-encoded image content
    """
    try:
        # Decode base64
        image_bytes = base64.b64decode(base64_content)

        # Check if compression is needed
        if len(image_bytes) <= max_size_bytes:
            logger.info(f"[IMAGE COMPRESS] Image already within limit: {len(image_bytes)} <= {max_size_bytes}")
            return base64_content

        # Compress
        compressed_bytes = compress_image(
            image_bytes,
            max_size_bytes=max_size_bytes,
            max_dimension=max_dimension,
            quality=quality,
        )

        # Encode back to base64
        return base64.b64encode(compressed_bytes).decode("utf-8")

    except Exception as e:
        logger.error(f"[IMAGE COMPRESS] Error processing base64 image: {e}")
        return base64_content
