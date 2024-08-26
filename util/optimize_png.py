import argparse
import os
from concurrent.futures import ProcessPoolExecutor
from PIL import Image, ImageOps
from tqdm import tqdm


def optimize_image(file_path, threshold):
    try:
        with Image.open(file_path) as img:
            # Convert to RGB if image has an alpha channel
            if img.mode == 'RGBA':
                img = img.convert('RGB')

            # Resize if larger than threshold
            if max(img.size) > threshold:
                img.thumbnail((threshold, threshold), Image.LANCZOS)

            # Apply additional optimizations
            img = ImageOps.posterize(img, 4)  # Reduce color palette
            img = ImageOps.autocontrast(img)  # Enhance contrast

            # Save with higher compression
            img.save(file_path, "PNG", optimize=True, quality=70, compress_level=9)
        return True
    except Exception as e:
        print(f"Error processing {file_path}: {str(e)}")
        return False


def process_images(input_folder, threshold):
    image_files = [
        os.path.join(root, file)
        for root, _, files in os.walk(input_folder)
        for file in files
        if file.lower().endswith(".png")
    ]

    with ProcessPoolExecutor() as executor:
        results = list(
            tqdm(
                executor.map(
                    optimize_image, image_files, [threshold] * len(image_files)
                ),
                total=len(image_files),
                desc="Optimizing images",
            )
        )

    successful = sum(results)
    print(f"Processed {successful} out of {len(image_files)} images successfully.")


def main():
    parser = argparse.ArgumentParser(description="Optimize PNG images in a folder")
    parser.add_argument(
        "input_folder", help="Path to the input folder containing PNG images"
    )
    parser.add_argument(
        "-t",
        "--threshold",
        type=int,
        default=1024,
        help="Max size for image resizing (default: 1024)",
    )
    args = parser.parse_args()

    if not os.path.isdir(args.input_folder):
        print(f"Error: {args.input_folder} is not a valid directory")
        return

    process_images(args.input_folder, args.threshold)


if __name__ == "__main__":
    main()