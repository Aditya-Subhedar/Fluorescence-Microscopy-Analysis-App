# Fluorescence Microscopy Analysis App

A professional-grade, desktop-based GUI tool developed to automate the preprocessing, manual correction, and quantification of multi-channel fluorescence microscopy images. 

This application provides a robust, efficient, and accurate alternative to manual image analysis workflows (e.g., ImageJ).

## Features

### Tab 1: CZI & Z-Stack Preprocessing
* **Robust File Support:** Robustly loads complex Zeiss CZI and multi-stack TIF files.
* **Intelligent Dimension Correction:** Automatically resolves Z-slice/Color-Channel misalignment traps and unstitched mosaic issues inherent in raw python microscopy readers.
* **Image Management:** Features a multi-image navigation bar supporting `askopenfilenames` (batch select) with keyboard arrow shortcut support (◀/▶) for streamlined workflow.
* **Visual Balancing:** Real-time composite view with integrated standard color mapping, interactive cropping tools, and side-by-side Contrast/Brightness optimization optimized for high-resolution displays.

### Tab 2: Automated & Manual Quantification
* **Automated Segmentation:** Utilizes adaptive and Otsu's method thresholding for accurate object detection.
* **Interactive Filtering:** Five real-time analysis sliders to filter objects by:
    1. **Hue Range:** Select specific fluorescent color bands.
    2. **Intensity (Min/Max):** Eliminate background noise and saturate signal.
    3. **Area Size (px):** Exclude objects outside relevant size ranges.
    4. **Circularity:** Avoid capturing elongated structures like nerve fibers, prioritizing cellular structures.
* **Precision Edit Tools:** Interactive canvas supporting:
    * **Manual Pencil/Eraser:** For pixel-perfect correction of automated segmentation.
    * **Undo/Redo:** Full state history support for edit rollback.
    * **Clear Drawings:** Instantly reset manual corrections.
* **Automatic Reporting:** Real-time display of **Fluorescent Area %** and **Cluster Counts**, with automated, structured export directly to Microsoft Excel.

## Setup & Installation

Ensure you have Python 3.8+ installed.

1.  **Clone this repository:**
    ```bash
    git clone [https://github.com/Aditya-Subhedar/Fluorescence-Microscopy-Analysis-App](https://github.com/your-username/fluorescence-microscopy-analysis-app.git)
    cd fluorescence-microscopy-analysis-app
    ```

2.  **Create and activate a virtual environment:**
    ```bash
    # Windows
    python -m venv venv
    venv\Scripts\activate
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

Place your raw CZI images into the `IHC input images` directory (this folder is untracked by Git). Launch the application by running:

```bash
python main_app.py