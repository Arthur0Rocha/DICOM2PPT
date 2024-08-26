import os
import argparse

import pydicom
import imageio
from pptx import Presentation
from pptx.util import Inches


def extract_images_from_dicom(dicom_seq_folder):
    images = [
        pydicom.dcmread(os.path.join(dicom_seq_folder, dicom_file))
        for dicom_file in os.listdir(dicom_seq_folder)
    ]
    images = [im.pixel_array for im in images if hasattr(im, "pixel_array")]
    maxes = max([im.max() for im in images]) if images else 1
    images = [
        im if maxes < 2**8 else im // (2**4) if maxes < 2**12 else im // 2**8 + 2**8
        for im in images
    ]
    return images


def add_image_slide(prs, image_file):
    slide_layout = prs.slide_layouts[5]  # Blank slide layout
    slide = prs.slides.add_slide(slide_layout)
    left = top = Inches(1)
    slide.shapes.add_picture(image_file, left, top, width=Inches(8.5), height=Inches(6))


def export_to_ppt(images_files, ppt_file):
    prs = Presentation()
    for image_file in images_files:
        add_image_slide(prs, image_file)
    prs.save(ppt_file)


def setup_parser():
    parser = argparse.ArgumentParser(
        prog="DICOM2PPT",
        description="Converts multiple DICOM files to images or GIFs and insert them into a PPT file",
        epilog="Usage: inputfolder outputfolder",
    )
    parser.add_argument("inputfolder")
    parser.add_argument("outputfolder")
    return parser


def main():
    parser = setup_parser()
    args = parser.parse_args()

    dicom_folder = args.inputfolder
    outpath = args.outputfolder

    gif_paths = []

    for sequence_folder in os.listdir(dicom_folder):
        print(f"Reading {sequence_folder}")
        infolder = os.path.join(dicom_folder, sequence_folder)
        sequence = extract_images_from_dicom(infolder)
        if not sequence:
            print("Skipping...")
            continue
        outgifpath = os.path.join(outpath, sequence_folder) + ".gif"
        try:
            imageio.mimsave(outgifpath, sequence)
        except Exception:
            imageio.mimsave(
                outgifpath, [im for im in sequence if im.shape == sequence[0].shape]
            )
            for i, im in enumerate(sequence):
                if im.shape != sequence[0].shape:
                    imageio.imwrite(f"{outgifpath}.{i}.png", im)
                    print(f"Bad shape: {i}")

        gif_paths.append(outgifpath)

    export_to_ppt(gif_paths, os.path.join(outpath, "presentation.pptx"))


if __name__ == "__main__":
    main()
