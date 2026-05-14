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

### 📍 Tab 3: Golgi Morphological Profiling (Sholl Analysis)
* **Neuronal Architecture Mapping:** Specialized 2D analytical suite optimized for tracking complex cellular extensions, dendrite branching networks, and dendritic spine concentrations.
* **Semi-Automated Sholl Calibration:** Interactive circle distance selection ($\mu m$) that automatically projects concentric bounding rings centered on the neuron's soma to calculate dendritic intersection densities at set radial distances.
* **Synced Duplex Viewing Engine:** Features a side-by-side presentation board splitting the raw, contrast-adjusted microscope frame from the high-contrast binary overlay stencil to trace morphology in real time.
* **Live Quantification Metrics:** Instant quantitative calculation tracking total detected spines, active spine frequencies within target ranges (e.g., $20\text{--}30\,\mu m$), and structural branch densities.
* **Session Logging:** Single-click tabular compilation exporting complete morphological tracking history arrays straight into structured Excel/CSV sheets for downstream figure generation.

### 📍 Tab 4: Figure Layer Merger & Annotation
* **Publication Figure Compositing:** Multi-layer alpha-channel blender that combines distinct custom-colored outline masks (e.g., separate channels for Red, Green, or Blue regions) into a single composite asset.
* **Microscopy Backdrop Superimposition:** Allows loading raw grayscale or multi-channel microscope files (`.tif`, `.jpeg`, `.png`) as the base layer, allowing users to overlay boundary vectors exactly on top of original cellular stain structures.
* **Flexible Backdrop Management:** Full support for transparent exports (for custom downstream manuscript assemblies) or solid background colors (e.g., crisp print-ready solid white or black backdrops).
* **Interactive Typography Overlays:** Built-in annotation suite supporting custom text, responsive font size scaling, and color configurations. 
* **Draggable Canvas Layout Elements:** Features a cursor drag-and-drop system to position text annotations anywhere on the canvas workspace. Text coordinates stay locked to high-resolution pixels to prevent pixelation upon export.

---

## 🛠️ Technical Architecture & Stack

* **GUI Engine:** Python `tkinter` + `ttk` widgets
* **Image Processing Engine:** OpenCV (`cv2`), `scikit-image` (`measure`), NumPy
* **I/O File Handlers:** Pillow (`PIL`), `czifile` image arrays
* **Data Pipelines:** `pandas`, `openpyxl`

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