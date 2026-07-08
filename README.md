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
* **Automated Segmentation:** Utilizes adaptive and Otsu's method thresholding for accurately segmenting fluoroscent regions. Auto detect default values can be changed according to user preferences.
* **Interactive Filtering:** Real-time analysis sliders to filter objects by Hue Range, Intensity (Min/Max), Area Size (px), and Morphology (isolating cells and fibres).
* **Batch Processing:** Multiple files can be loaded at once with each preserving its thresholding parameters during switching and upto 20 states for undo and redo operations. Supports .ome.tif, tiff, jpeg, jpg, png files.
* **Pre-Fetch Caching Engine:** Powered by an asynchronous background thread pool (`ThreadPoolExecutor`) that pre-loads adjacent images into RAM, enabling instant image switching with zero file-loading lag.
* **Smart Window Reset:** Canvas interface equipped with automatic viewport centering on file load and instant 1.0x unzoom snap execution on canvas double-click.
* **Precision Edit Tools:** Interactive canvas supporting manual pencil/eraser corrections with fast ellipse drawing tool, full state history (Undo/Redo), and automated exports to Microsoft Excel.
* **Hardware Calibration Tracking:** Integrated spatial metadata extraction subsystem that decodes embedded OME-TIFF tags and native Zeiss CZI hardware XML parameters to calculate absolute fluorescent surface areas in micrometers (um2).
* **Real-Time Fluoroscence Data:** Fluoroscence data is displayed in real time at the bottom with fields "Fluoroscent Area (%)", "Fluoroscent Area (Absolute (um2))", Cluster Count.
* **Saves Data, Presets, Contours and Overlayed Images:** Quantitative fluoroscent data is for batch of multiple images can be stored in CSV and EXCEL format. Thresholding parameters are stored in a JSON file named "cytoquant_presets.json" which can be accessed through "apply presets button". Contours can be saved in PNG format and superimposed on images. Images with overlayed contours (including manual drawings)can be saved in JPEG, PNG, TIFF (un-compressed)formats.

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

### 📍 Tab 6: Open Field Test (OFT) Tracking
* **Adaptive Motion Tracking:** Utilizes a robust background subtraction engine (MOG2) paired with morphological filtering to continuously isolate and track the subject's center of mass across complex lighting environments.
* **Real-World Spatial Calibration:** Features an interactive 4-point perspective transform matrix (Homography) that maps raw pixel coordinates to physical arena dimensions for both square and circle OFT chambers, distance (length and diameter) of both inner and outer zones of square and circle OFT chambers can be set manually and their geometry adapts automatically ensuring absolute distance accuracy regardless of camera angle distortion.
* **Kinematic Noise Filtering:** Implements inertial smoothing and velocity deadzones to prevent false distance spikes ("teleportation" artifacts) when a user manually repositions the tracking bounding box during a trial. The tracker inertia and minimum deadzone distance is also adjustable in real time using sliders on the UI.
* **Live Behavioral Analytics:** Real-time quantitative computation of total trajectory distance (cm), crossover count, and central zone spatial preference (time spent in center).
* **Fluid Media Controls:** Integrated video playback system featuring global keyboard shortcut bindings (Spacebar for Play/Pause, ◄/► for precise 2-second seeking) safely decoupled from UI widget focus.
* **Comprehensive Data Export:** Single-click production exports generating both structured quantitative metric tables (`.csv`) and high-contrast spatial trajectory maps (`.png`) for publication figures.

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

Place your raw CZI/OME-TIFF images into any directory. Launch the application by running:

```bash
python main_app.py


## 📦 Compilation to Standalone Executable

To compile the application into a standalone standalone desktop application (`.exe`) within your virtual environment, run the python module execution flag command:

```bash
python -m PyInstaller --onefile --windowed --icon=logo.ico --name="CytoQuant_V19" --hidden-import=czifile --hidden-import=pylibCZIrw --hidden-import=tifffile --hidden-import=skimage main_app.py
```
The finished production package will be placed inside the generated `dist/` directory.