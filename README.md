# Fluorescence Microscopy Analysis App

An imaging software providing end to end functionality from raw microsocpe image to publication ready data and panel generation. A desktop-based GUI tool developed to automate the preprocessing, manual correction, quantification of multi-channel fluorescence microscopy images, image annotation and panel generation. 

This application provides a robust, efficient, and accurate alternative to manual image analysis workflows (e.g., ImageJ).

## Features

### 📍 Tab 1: Microscope Image Preprocessing
* **Robust File Support:** Robustly loads OME file formats such as OME-TIFF, TIF, CZI (Zeiss), LIF (Leica), ND2 (Nikon), OIB (Olympus) files.
* **Intelligent Dimension Correction:** Auto range expansion, Multiple Z-stack collapsing methods, Color-Channel pseudocoloring and unstitched mosaic issues inherent in raw python microscopy readers.
* **Image Management:** Features a multi-image navigation bar supporting `askopenfilenames` (batch select) with keyboard arrow shortcut support (←/→), traverse z-stack using up-down arrow keys (↑,↓) for streamlined workflow.
* **Fluid Viewport Control:** Lag-free interactive canvas supporting precise scroll-wheel zooming and smooth click-and-drag panning.
* **Fidelity Optimization:** Adaptive hybrid rendering engine that automatically switches processing workloads to match the viewport size, delivering crystal-clear full-resolution details during deep zoom passes without lagging.

### 📍 Tab 2: Automated & Manual Quantification
* **Automated Segmentation:** Utilizes adaptive and Otsu's method thresholding for accurately segmenting fluorescent regions. Auto detect default values can be changed according to user preferences.
* **Interactive Filtering:** Real-time analysis sliders to filter objects by Hue Range, Intensity (Min/Max), Area Size (px), and Morphology (isolating cells and fibres).
* **Batch Processing:** Multiple files can be loaded at once with each preserving its thresholding parameters during switching and up to 20 states for undo and redo operations. Supports .ome.tif, tiff, jpeg, jpg, png files.
* **Pre-Fetch Caching Engine:** Powered by an asynchronous background thread pool (`ThreadPoolExecutor`) that pre-loads adjacent images into RAM, enabling instant image switching with zero file-loading lag.
* **Smart Window Reset:** Canvas interface equipped with automatic viewport centering on file load and instant 1.0x unzoom snap execution on canvas double-click.
* **Precision Edit Tools:** Interactive canvas supporting manual pencil/eraser corrections with a fast ellipse drawing tool, a **dynamic "Lasso Fill" algorithm** that instantly converts hollow hand-drawn contours into solid analytical regions, and full state history (Undo/Redo).
* **Hardware Calibration Tracking:** Integrated spatial metadata extraction subsystem that decodes embedded OME-TIFF tags and native Zeiss CZI hardware XML parameters to calculate absolute fluorescent surface areas in micrometers (um2).
* **Real-Time Fluorescence Data & UI Toggle:** Fluorescence data is displayed in real time at the bottom with fields "Fluorescent Area (%)", "Fluorescent Area (Absolute (um2))", and Cluster Count. A **toggleable ROI index** is displayed on the preview (which dynamically accounts for manual drawing/erasing) allowing users to easily hide numbers for clean visual inspection.
* **Single-ROI Multi-Parametric Profiling:** The engine extracts granular metrics for every isolated shape, calculating individual Mean Gray Value, Integrated Density, Area Fraction, Circularity, Eccentricity, and a vectorized Nearest-Neighbor proximity matrix. 
* **Saves Data, Presets, Contours and Overlaid Images:** Quantitative fluorescent data for a batch of multiple images is compiled into **custom multi-sheet EXCEL workbooks** (featuring a "Global Batch Overview" alongside individual image data sheets) and CSV formats. Thresholding parameters are stored in a JSON file named "cytoquant_presets.json" which can be accessed through the "apply presets" button. Contours can be saved in PNG format and superimposed on images. Images with overlaid contours (including manual drawings) can be saved in JPEG, PNG, and TIFF (un-compressed) formats.

### 📍 Tab 3: Figure Layer Merger & Annotation
* **Publication Figure Compositing:** Multi-layer alpha-channel blender that combines distinct custom-colored outline masks into a single composite asset.
* **Microscopy Backdrop Superimposition:** Allows loading raw grayscale or multi-channel microscope files (`.tif`, `.jpeg`, `.png`) as the base layer, allowing users to overlay boundary vectors exactly on top of original cellular stain structures.
* **Flexible Backdrop Management:** Full support for transparent exports (for custom downstream manuscript assemblies) or solid background colors (e.g., crisp print-ready solid white or black backdrops).
* **Interactive Typography Overlays:** Built-in annotation suite supporting custom text, responsive font size scaling, and color configurations. 
* **Draggable Canvas Layout Elements:** Features a cursor drag-and-drop system to position text annotations anywhere on the canvas workspace. Text coordinates stay locked to high-resolution pixels to prevent pixelation upon image rotation and image export.

### 📍 Tab 4: Multi-Image Publication Panel Creator
* **Interactive Matrix Generation:** Generates a clean, structural layout matrix based on custom row and column configurations with an explicit image allocation wizard built into every cell.
* **Geometric Aspect Ratio Presets:** Configures cell bounds using standard scientific presentation formats (Square 1:1, Landscape, Portrait) with adjustable inter-image pixel gaps.
* **On-Canvas Image Previews:** Loads raw image assets directly into targeted grid placeholders, instantly rendering downscaled thumbnails alongside automated aspect ratio mismatch safety warnings.
* **Automated Subpanel Indexing:** Includes a sequential lettering engine (e.g., A., B., C...) that positions crisp, high-contrast labels across custom layouts (Inside vs. Outside Top-Left positions).
* **High-Fidelity Canvas Compositing:** Assembles the high-resolution final panel independently from the UI layout, using Lanczos interpolation resampling to protect original image clarity while vector-mapping multi-axis text labels and margin captions.
* **Lossless Resolution Export:** Supports single-click production exports into publication-grade formats (`.png`, `.jpeg`, `.tiff`) optimized for direct integration into PowerPoint presentations or manuscript documents.

### 📍 Tab 5: Golgi Morphological Profiling (Sholl Analysis)
* **Neuronal Architecture Mapping:** Specialized 2D analytical suite optimized for tracking complex cellular extensions, dendrite branching networks, and dendritic spine concentrations.
* **Semi-Automated Sholl Calibration:** Interactive circle distance selection ($\mu m$) that automatically projects concentric bounding rings centered on the neuron's soma to calculate dendritic intersection densities at set radial distances.
* **Synced Duplex Viewing Engine:** Features a side-by-side presentation board splitting the raw, contrast-adjusted microscope frame from the high-contrast binary overlay stencil to trace morphology in real time.
* **Live Quantification Metrics:** Instant quantitative calculation tracking total detected spines, active spine frequencies within target ranges (e.g., $20\text{--}30\,\mu m$), and structural branch densities.
* **Session Logging:** Single-click tabular compilation exporting complete morphological tracking history arrays straight into structured Excel/CSV sheets for downstream figure generation.

### 📍 Tab 6: Automated Statistical Visualization & Graph Generation
* **Flexible Plot Formats:** Ingests raw replicate data matrices to generate aggregated Grouped Bar charts, Box plots, or Violin plots calculating central tendencies mapped against user-selected Standard Deviation (SD) or Standard Error of the Mean (SEM) dispersion metrics.
* **Replicate Dispersion Jitter:** Transparently overlays raw data points over bar charts using an adjustable, randomized Gaussian spatial jitter to prevent structural occlusion within high-density replicate clusters.
* **Inferential Statistical Backend:** Computes automated parametric or non-parametric pairwise comparisons (Independent t-tests, Mann-Whitney U tests) and global variance profiles (One-Way ANOVA) directly within the interface against an alpha threshold (a = 0.05).
* **Dynamic Significance Routing:** Automatically parses the maximum local data heights (y_max) to programmatically route adaptive, stacked significance brackets accented with standard academic nomenclature (*, **, ***) while applying built-in collision prevention.
* **Granular Aesthetic Controls:** Exposes absolute custom control over minor presentation properties including independent bar widths, intra-group gap spacing, custom hexadecimal group colors, structural fill hatch patterns, and explicit Y-axis boundary scaling.
* **Dynamic Layout Overlap Protection:** Integrates an adaptive canvas margin compression engine that dynamically shifts the subplot grid boundaries when assigning an external right-side legend, ensuring long group string titles never overlap with ticks or significance brackets.
* **Lossless Figure Rendering:** Fully decoupled from UI viewport constraints to export crisp, publication-grade graph figures directly into 300 DPI loss-less `.png`, `.jpeg`, or `.tiff` formats.
---

## 🛠️ Technical Architecture & Stack

* **GUI Engine:** Python `tkinter` + `ttk` widgets
* **Image Processing Engine:** OpenCV (`cv2`), `scikit-image` (`filters`, `measure`), NumPy
* **I/O File Handlers:** Pillow (`PIL`), `czifile`, `tifffile` *(and `pylibCZIrw` if integrated)*
* **Concurrency Engine:** `threading`, `concurrent.futures`
* **Data Pipelines:** `pandas`, `csv`, `openpyxl`

## Setup & Installation

Ensure you have Python 3.8+ installed.

1.  **Clone this repository:**
    ```bash
    git clone [https://github.com/Aditya-Subhedar/Fluorescence-Microscopy-Analysis-App.git](https://github.com/Aditya-Subhedar/Fluorescence-Microscopy-Analysis-App.git)
    cd Fluorescence-Microscopy-Analysis-App
    ```

2.  **Create and activate a virtual environment:**
    ```bash
    # For Windows:
    python -m venv venv
    venv\Scripts\activate

    # For macOS and Linux:
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

Place your raw CZI/OME-TIFF images into any directory. Launch the application by running:

```bash
python main_app.py


## 📦 Compilation to Standalone Executable

To compile the application into a standalone standalone desktop application (`.exe`) within your virtual environment, run the python module execution flag command:

```bash
python -m PyInstaller --onefile --windowed --icon=logo.ico --name="CytoQuant" --hidden-import=czifile --hidden-import=pylibCZIrw --hidden-import=tifffile --hidden-import=skimage main_app.py
```
The finished production package will be placed inside the generated `dist/` directory.