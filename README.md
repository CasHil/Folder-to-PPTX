# Images to PowerPoint Presentation

This script creates a PowerPoint presentation from images in a directory. The images are added to the presentation as centered slides.

## Requirements

This script requires Python and the following Python packages:

- `python-pptx`
- `Pillow`

You can install these packages using pip:

```
pip install python-pptx Pillow
```

or run the following command:

```
pip install -r requirements.txt
```

The script is written in Python 3.12.1.

## Usage
You can run the script from the command line like this:

```
python images-to-pptx.py --directory [/path/to/images] --filename [presentation.pptx]
```

- `--directory`: The directory containing the images. This should be a path to a directory on your system. If not provided, the script will use the current working directory. The directory is traversed recursively, so all images in subdirectories will be added to the presentation.
- `--filename`: The filename for the output presentation. If not provided, the script will use 'presentation.pptx' as the default filename.
The script will traverse the specified directory (or the current directory if no directory is specified) and add all images to the presentation. The supported image formats are PNG, JPG, JPEG, GIF, BMP, TIFF, and WMF.

Both arguments are optional.