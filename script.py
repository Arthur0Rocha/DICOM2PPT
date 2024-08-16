import pydicom
import numpy as np
from PIL import Image


def extract_images_from_dicom(dicom_file):
    ds = pydicom.dcmread(dicom_file)
    images = []

    # Check if the DICOM file contains multiple frames
    if hasattr(ds, "NumberOfFrames") and ds.NumberOfFrames > 1:
        for i in range(ds.NumberOfFrames):
            image = ds.pixel_array[i]  # Extract pixel array for each frame
            images.append(image)
    else:
        images.append(ds.pixel_array)  # Single frame

    return images


# Example usage
dicom_file = "path/to/your/dicom/file.dcm"
images = extract_images_from_dicom(dicom_file)

# ----------------------------------------------------------- #

import imageio


def create_gif(images, gif_file):
    with imageio.get_writer(gif_file, mode="I") as writer:
        for image in images:
            writer.append_data(np.array(image))


# Example usage
gif_file = "output.gif"
create_gif(images, gif_file)

# ----------------------------------------------------------- #

import imageio


def create_video(images, video_file, fps=15):
    with imageio.get_writer(video_file, mode="I", fps=fps) as writer:
        for image in images:
            writer.append_data(np.array(image))


# Example usage
video_file = "output.mp4"
create_video(images, video_file)

# ----------------------------------------------------------- #

from pptx import Presentation
from pptx.util import Inches


def add_image_slide(prs, image_file):
    slide_layout = prs.slide_layouts[5]  # Blank slide layout
    slide = prs.slides.add_slide(slide_layout)
    left = top = Inches(1)
    pic = slide.shapes.add_picture(
        image_file, left, top, width=Inches(8.5), height=Inches(6)
    )


def export_to_ppt(images_files, ppt_file):
    prs = Presentation()
    for image_file in images_files:
        add_image_slide(prs, image_file)
    prs.save(ppt_file)


# Example usage
ppt_file = "output.pptx"
export_to_ppt([gif_file, video_file], ppt_file)
