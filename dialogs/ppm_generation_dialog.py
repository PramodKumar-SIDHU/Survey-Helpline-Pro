from qgis.PyQt.QtWidgets import QDialog, QFileDialog, QMessageBox
from .ppm_generation_dialog_ui import Ui_PropertyParcelMapsGeneration
from qgis.core import QgsProject, QgsWkbTypes, QgsVectorLayer, QgsVectorFileWriter, QgsCoordinateTransformContext, QgsFeature,QgsField,QgsVectorDataProvider,QgsLayerTreeLayer,QgsSpatialIndex, QgsLayerTreeGroup
from qgis.PyQt.QtCore import QThread, pyqtSignal, QVariant
import time, os, shutil
import processing
from qgis.PyQt import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QListView, QMessageBox
from PyQt5.QtGui import QFontDatabase
from qgis.PyQt.QtXml import QDomDocument
from PyQt5.QtCore import QTimer, QEventLoop
from qgis.core import (
    QgsProject,
    QgsVectorLayer,
    QgsWkbTypes,
    QgsFeature,
    QgsField,
    QgsVectorDataProvider,
    QgsExpression,
    QgsSymbol,
    QgsFillSymbol,
    QgsLineSymbol,
    QgsMarkerSymbol,
    QgsRuleBasedRenderer,
    QgsPalLayerSettings,
    QgsTextFormat,
    QgsTextBufferSettings,
    QgsVectorLayerSimpleLabeling,
    QgsTextBufferSettings,
    QgsProperty,
    QgsFeatureRequest,
    QgsGeometry,
    QgsVectorFileWriter,
    QgsRuleBasedLabeling,
    QgsTextBackgroundSettings,
    QgsPrintLayout,
    QgsLayoutAtlas,
    QgsLayoutItemLabel,
    QgsLayoutItemHtml,
    QgsReadWriteContext,
    QgsPropertyCollection,
    QgsUnitTypes,
    QgsCoordinateReferenceSystem,
    QgsCoordinateTransform,
    QgsMapSettings,
    QgsMapThemeCollection,
    QgsApplication
)
from collections import defaultdict

class PPMGenerationDialog(QDialog, Ui_PropertyParcelMapsGeneration):
    def __init__(self, iface, parent=None):
        super(PPMGenerationDialog, self).__init__(parent)
        self.iface = iface
        self.setupUi(self)
        self.init_ui()

    def init_ui(self):

        # ---------------------------- #
        #  District Dropdown
        # ---------------------------- #
        self.cb_district.clear()
        self.cb_district.addItems([
            "-- Select District --",
            "Alluri Sitarama Raju", "Anakapalli", "Anantapur", "Annamayya",
            "Bapatla", "Chittoor", "Dr. B.R. Ambedkar Konaseema",
            "East Godavari", "Eluru", "Guntur", "Kakinada", "Krishna",
            "Kurnool", "Nandyal", "NTR", "Palnadu", "Parvathipuram Manyam",
            "Prakasam", "Sri Potti Sriramulu Nellore", "Sri Satya Sai",
            "Srikakulam", "Tirupati", "Visakhapatnam", "Vizianagaram",
            "West Godavari", "YSR Kadapa"
        ])

        # ---------------------------- #
        #  Populate Shapefile Dropdowns
        # ---------------------------- #
        self.populate_shapefile_dropdowns()

        # Update parcel columns when parcel layer changes
        self.cb_parcel_shp.currentIndexChanged.connect(self.update_parcel_columns)

        self.cb_builtup_shp.currentIndexChanged.connect(self.update_builtup_columns)

        # Freeze builtup column dropdown initially
        self.cb_builtup_ppn.setEnabled(False)

        # ---------------------------- #
        #  Browse Buttons
        # ---------------------------- #
        self.btn_browse_excel.clicked.connect(self.browse_excel)
        self.btn_browse_output.clicked.connect(self.browse_output)

        # ---------------------------- #
        #  Enable / Disable Excel Browse
        # ---------------------------- #
        self.chk_ppms_92_notice.toggled.connect(self.on_ppms_92_toggled)
        # self.chk_ppms_92_notice.stateChanged.connect(self.toggle_excel)

        # Initially disable the Excel browse button
        self.btn_browse_excel.setEnabled(False)

        # ---------------------------- #
        #  Dialog Buttons (Reset / Close)
        # ---------------------------- #
        self.buttonBox_2.rejected.connect(self.close)
        self.buttonBox_2.clicked.connect(self.on_buttonbox_clicked)

        # ---------------------------- #
        #  Excel Download
        # ---------------------------- #
        self.excel_download.clicked.connect(self.download_excel_template)

        # ---------------------------- #
        #  Run Button
        # ---------------------------- #
        self.btn_run.clicked.connect(self.run_process)
        self.current_step = 0
        self.total_steps = 1

    def download_excel_template(self):
        """
        Let user select a folder and copy the 9(2)_Attributes_Excel file to it
        """

        # Ask user to select folder
        dest_folder = QFileDialog.getExistingDirectory(
            self,
            "Select Folder to Save Excel",
            "",
            QFileDialog.ShowDirsOnly
        )

        if not dest_folder:
            return  # user cancelled

        # Path of this plugin folder
        plugin_dir = os.path.dirname(__file__)

        # Source Excel file (inside plugin folder)
        src_excel = os.path.join(
            plugin_dir,
            "9(2)_Attributes_Excel.xlsx"
        )

        if not os.path.exists(src_excel):
            QMessageBox.critical(
                self,
                "File Not Found",
                "9(2)_Attributes_Excel.xlsx not found in plugin folder."
            )
            return

        # Destination path
        dest_excel = os.path.join(
            dest_folder,
            "9(2)_Attributes_Excel.xlsx"
        )

        try:
            shutil.copy(src_excel, dest_excel)

            QMessageBox.information(
                self,
                "Download Complete",
                f"Excel file saved successfully:\n{dest_excel}"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to copy Excel file.\n\n{e}"
            )

    def on_buttonbox_clicked(self, button):
        role = self.buttonBox_2.buttonRole(button)

        if role == QtWidgets.QDialogButtonBox.ResetRole:
            # Switch to Parameters tab
            index = self.tabWidget.indexOf(self.Parameters)
            if index != -1:
                self.tabWidget.setCurrentIndex(index)

    def populate_shapefile_dropdowns(self):
        """Populate shapefile dropdowns with all polygon vector layers in project."""
        self.cb_parcel_shp.clear()
        self.cb_builtup_shp.clear()

        layers = QgsProject.instance().mapLayers().values()

        for lyr in layers:
            if lyr.type() == lyr.VectorLayer:
                # Optional: Only polygon layers
                if lyr.geometryType() == QgsWkbTypes.PolygonGeometry:
                    self.cb_parcel_shp.addItem(lyr.name(), lyr.id())
                    self.cb_builtup_shp.addItem(lyr.name(), lyr.id())

    def update_parcel_columns(self):
        """Update column dropdown based on selected parcel shapefile."""
        self.cb_initial_ppn.blockSignals(True)   # Prevent unwanted triggers
        self.cb_initial_ppn.clear()

        # Add placeholder
        self.cb_initial_ppn.addItem("-- Select Column --", None)

        lyr_id = self.cb_parcel_shp.currentData()
        if not lyr_id:
            self.cb_initial_ppn.blockSignals(False)
            return

        layer = QgsProject.instance().mapLayer(lyr_id)
        if layer is None:
            self.cb_initial_ppn.blockSignals(False)
            return

        for field in layer.fields():
            self.cb_initial_ppn.addItem(field.name(), field.name())

        self.cb_initial_ppn.blockSignals(False)

        # Connect validator (only once)
        try:
            self.cb_initial_ppn.currentIndexChanged.disconnect(self.validate_ppn_column)
        except:
            pass
        self.cb_initial_ppn.currentIndexChanged.connect(self.validate_ppn_column)

    def validate_ppn_column(self):
        """Validate the selected PPN column for blanks and duplicates."""

        selected_col = self.cb_initial_ppn.currentData()
        if selected_col is None:
            return  # Placeholder selected

        lyr_id = self.cb_parcel_shp.currentData()
        layer = QgsProject.instance().mapLayer(lyr_id)
        if layer is None:
            return

        blanks = 0
        values = []
        duplicates = {}

        for f in layer.getFeatures():
            val = f[selected_col]

            if val is None or str(val).strip() == "":
                blanks += 1
            else:
                if val in values:
                    if val not in duplicates:
                        duplicates[val] = 1
                    duplicates[val] += 1
                values.append(val)

        # If no issues → return silently
        if blanks == 0 and len(duplicates) == 0:
            return

        # ----------------------------------- #
        #  Build Error Message
        # ----------------------------------- #
        msg = "<b>Validation Errors in Column:</b> <br><br>"

        if blanks > 0:
            msg += f"• <b>{blanks}</b> blank or null values found.<br><br>"

        if len(duplicates) > 0:
            msg += "<b>Duplicate values found:</b><br>"
            for v, count in duplicates.items():
                msg += f"– {v} (repeated {count} times)<br>"

        # ----------------------------------- #
        #  Show Popup Warning
        # ----------------------------------- #
        QMessageBox.warning(self, "Invalid PPN / Chalta No. Column", msg)

        # ----------------------------------- #
        #  RESET THE DROPDOWN
        # ----------------------------------- #
        self.cb_initial_ppn.blockSignals(True)
        self.cb_initial_ppn.setCurrentIndex(0)   # Reset to "-- Select Column --"
        self.cb_initial_ppn.blockSignals(False)

    def on_ppms_92_toggled(self, checked):
        """
        If 9(2) Notice is checked, disable & uncheck other options.
        """
        if checked:
            # Uncheck others
            self.chk_village_map.setChecked(False)
            self.chk_initial_ppms.setChecked(False)

            # (Optional but recommended) Disable them
            self.chk_village_map.setEnabled(False)
            self.chk_initial_ppms.setEnabled(False)
        
            # Enable Excel browse
            self.btn_browse_excel.setEnabled(True)
            self.cb_builtup_ppn.setEnabled(True)
            # self.cb_builtup_ppn.currentIndexChanged.connect(self.update_builtup_columns)

            # Populate columns immediately
            self.update_builtup_columns()
            
        else:
            # Re-enable when unchecked
            self.chk_village_map.setEnabled(True)
            self.chk_initial_ppms.setEnabled(True)

            # Disable Excel browse
            self.btn_browse_excel.setEnabled(False)
            # self.cb_builtup_ppn.setEnabled(False)

            # Freeze & clear builtup column dropdown
            self.cb_builtup_ppn.blockSignals(True)
            self.cb_builtup_ppn.clear()
            self.cb_builtup_ppn.addItem("-- Select Column --", None)
            self.cb_builtup_ppn.setEnabled(False)
            self.cb_builtup_ppn.blockSignals(False)
            
    def update_builtup_columns(self):
        """Update column dropdown based on selected builtup shapefile."""
        self.cb_builtup_ppn.blockSignals(True)   # Prevent unwanted triggers
        self.cb_builtup_ppn.clear()

        # Add placeholder
        self.cb_builtup_ppn.addItem("-- Select Column --", None)

        lyr_id = self.cb_builtup_shp.currentData()
        if not lyr_id:
            self.cb_builtup_ppn.blockSignals(False)
            return

        layer = QgsProject.instance().mapLayer(lyr_id)
        if layer is None:
            self.cb_builtup_ppn.blockSignals(False)
            return

        for field in layer.fields():
            self.cb_builtup_ppn.addItem(field.name(), field.name())

        self.cb_builtup_ppn.blockSignals(False)

        # # Connect validator (only once)
        # try:
        #     self.cb_builtup_ppn.currentIndexChanged.disconnect(self.validate_ppn_column)
        # except:
        #     pass
        # self.cb_builtup_ppn.currentIndexChanged.connect(self.validate_ppn_column)
        
    def browse_excel(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Excel file", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.le_excel_path.setText(file)

    def browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.le_output_path.setText(folder)


    def append_log(self, message, log_type="info"):
        """Append message to QTextEdit with color coding"""
        color = {
            "info": "mediumvioletred",
            "success": "seagreen",
            "error": "crimson"
        }.get(log_type, "black")

        self.te_log.append(f'<span style="color:{color}">{message}</span>')
        self.te_log.verticalScrollBar().setValue(self.te_log.verticalScrollBar().maximum())

    def safe_export_layer(self, layer: QgsVectorLayer, output_path: str, layer_name: str):
        """
        Safely exports a vector layer and reloads it into project.

        Returns:
            QgsVectorLayer or None
        """

        if layer is None or not layer.isValid():
            self.append_log(f"❌ Layer '{layer_name}' is invalid or not found.", "error")
            return None

        try:
            err, msg = QgsVectorFileWriter.writeAsVectorFormat(
                layer,
                output_path,
                "UTF-8",
                layer.crs(),
                "GPKG"
            )

            if err != QgsVectorFileWriter.NoError:
                self.append_log(f"⚠️ Failed to export '{layer_name}': {msg}", "error")
                return None

            self.append_log(f"✅ Exported: {output_path}", "success")
            new_layer=self.readding_layer(layer_name, output_path)
            return new_layer
            # # Reload exported layer
            # new_layer = QgsVectorLayer(output_path, layer_name, "ogr")
            # if new_layer.isValid():
            #     QgsProject.instance().addMapLayer(new_layer)
            #     self.append_log(f"🔄 Reloaded layer: {layer_name}", "info")
            #     return new_layer
            # else:
            #     self.append_log(f"⚠️ Failed to reload '{layer_name}'", "error")
            #     return None

        except Exception as e:
            self.append_log(f"❌ Exception exporting '{layer_name}': {e}", "error")
            return None

    def readding_layer(self, layer_name, output_path):
        # Reload exported layer
        new_layer = QgsVectorLayer(output_path, layer_name, "ogr")
        if new_layer.isValid():
            QgsProject.instance().addMapLayer(new_layer)
            self.append_log(f"🔄 Reloaded layer: {layer_name}", "info")
            return new_layer
        else:
            self.append_log(f"⚠️ Failed to reload '{layer_name}'", "error")
            return None
    
    def add_fields_safe(self, layer, fields_dict):
        if not layer or not layer.isValid():
            self.append_log("❌ Invalid layer for field creation.", "error")
            return False

        if not layer.isEditable():
            layer.startEditing()

        provider = layer.dataProvider()
        new_fields = []

        for name, qtype in fields_dict.items():
            if layer.fields().indexFromName(name) == -1:
                new_fields.append(QgsField(name, qtype))

        if new_fields:
            provider.addAttributes(new_fields)
            layer.updateFields()
            self.append_log("📝 Fields created successfully.", "success")

        return True

    def update_attributes_safe(self, layer, updates):
        if not layer or not layer.isValid():
            self.append_log("❌ Invalid layer for attribute update.", "error")
            return False

        if not layer.isEditable():
            layer.startEditing()

        provider = layer.dataProvider()
        value_map = {}

        for fid, attrs in updates:
            field_index_map = {layer.fields().indexFromName(k): v for k, v in attrs.items()}
            value_map[fid] = field_index_map

        provider.changeAttributeValues(value_map)
        layer.updateFields()

        self.append_log("🔄 Attribute values updated.", "success")
        return True

    def commit_layer_safe(self, layer):
        if layer.isEditable():
            if layer.commitChanges():
                self.append_log("💾 Layer committed successfully.", "success")
                return True
            else:
                layer.rollBack()
                self.append_log("❌ Commit failed! Changes rolled back.", "error")
                return False
        return True

    def get_layer_by_name(self, name):
        """Return QgsVectorLayer object by dropdown name"""
        for layer in QgsProject.instance().mapLayers().values():
            if layer.name() == name:
                return layer
        return None

    def apply_qml_style(self, layer, qml_path):
        """
        Apply QML style to an already-loaded QgsVectorLayer.
        Handles QGIS false error-message bug.
        """
        try:
            if not os.path.exists(qml_path):
                self.append_log(f"⚠️ QML not found: {qml_path}","error")
                return

            ok, error_message = layer.loadNamedStyle(qml_path)

            # QGIS sometimes returns (True, "True")
            if ok or str(error_message).strip() == "True":
                QgsApplication.processEvents()
                layer.triggerRepaint()
                self.append_log(
                    f"🎨 Style applied from {os.path.basename(qml_path)} → {layer.name()}", "success")
                QgsApplication.processEvents()
            else:
                self.append_log(
                    f"⚠️ QML load issue for {layer.name()}: {error_message}", "info")

        except Exception as e:
            self.append_log(
                f"❌ Styling error for {layer.name()}: {str(e)}", "info")

    def create_theme(self, theme_name):
        try:
            project = QgsProject.instance()
            theme_collection = project.mapThemeCollection()
            root = project.layerTreeRoot()

            # Get REAL model (not QSortFilterProxyModel)
            proxy = self.iface.layerTreeView().model()
            real_model = proxy.sourceModel()

            # Create theme from current state (visibility + styles)
            theme_record = QgsMapThemeCollection.createThemeFromCurrentState(
                root,
                real_model
            )

            # Insert theme to collection
            theme_collection.insert(theme_name, theme_record)

            # Apply theme
            theme_collection.applyTheme(theme_name, root, real_model)

            self.iface.mapCanvas().refresh()

            self.append_log(f"🎨 Map theme '{theme_name}' created successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error creating map theme: {e}", "error")

    def calculate_total_area_fields(self, layer, m2_field="Ar_M2", ya2_field="Ar_Ya2",
                                    total_m2="ToT_Ar_M2", total_ya2="ToT_Ar_Y2"):
        """Calculate total area from Ar_M2 and Ar_Ya2 and store the SAME total in every feature."""
        
        if not layer.isEditable():
            layer.startEditing()

        provider = layer.dataProvider()
        fields = layer.fields().names()

        # Add new fields if missing
        fields_to_add = []
        if total_m2 not in fields:
            fields_to_add.append(QgsField(total_m2, QVariant.Double, len=20, prec=3))
        if total_ya2 not in fields:
            fields_to_add.append(QgsField(total_ya2, QVariant.Double, len=20, prec=3))
        if fields_to_add:
            provider.addAttributes(fields_to_add)
            layer.updateFields()

        # Indexes
        idx_tot_m2 = layer.fields().indexFromName(total_m2)
        idx_tot_ya2 = layer.fields().indexFromName(total_ya2)
        idx_m2 = layer.fields().indexFromName(m2_field)
        idx_ya2 = layer.fields().indexFromName(ya2_field)

        # ---- Compute totals ----
        total_area_m2 = 0
        total_area_ya2 = 0

        for feat in layer.getFeatures():
            total_area_m2 += float(feat[m2_field])
            # QgsApplication.processEvents()
            total_area_ya2 += float(feat[ya2_field])

        total_area_m2 = round(total_area_m2, 3)
        total_area_ya2 = round(total_area_ya2, 3)

        # ---- Apply totals to all features ----
        updates = {}
        for feat in layer.getFeatures():
            updates[feat.id()] = {
                idx_tot_m2: total_area_m2,
                idx_tot_ya2: total_area_ya2
            }


        provider.changeAttributeValues(updates)
        layer.commitChanges()
        layer.triggerRepaint()
        self.append_log("📐 Total Area Calculations completed.", "success")

    def open_layout(self, dest_qpt, parcel_layer, supproting_file):
        """Called in GUI thread — SAFE to use iface here."""
        try:

            project = QgsProject.instance()
            layout_mgr = project.layoutManager()

            layout = QgsPrintLayout(project)
            layout.initializeDefaults()

            doc = QDomDocument()
            with open(dest_qpt, "r", encoding="utf-8") as f:
                doc.setContent(f.read())

            context = QgsReadWriteContext()
            layout.loadFromTemplate(doc, context)

            layout.setName("INITIAL_PPM_A4_TEMPLATE_NEW") # IMPORTANT — set layout name AFTER loading template

            layout_mgr.addLayout(layout)
            
            # # Ensure layers are committed
            # parcel_layer.commitChanges()
            # supproting_file.commitChanges()
            # QgsApplication.processEvents()

            # Force provider + renderer + labeling refresh
            parcel_layer.triggerRepaint()
            supproting_file.triggerRepaint()
            QgsApplication.processEvents()

            # self.iface.mapCanvas().refreshAllLayers()
            # QgsApplication.processEvents()
            # self.iface.mapCanvas().refresh()
            # QgsApplication.processEvents()

            # Set the atlas coverage layer if layout has atlas
            atlas = layout.atlas()
            if atlas.enabled():
                atlas.setCoverageLayer(parcel_layer)
                self.append_log(f"🗺️ Atlas coverage layer set to: {parcel_layer.name()}", "success")

            # Open layout designer
            self.iface.openLayoutDesigner(layout)
            self.append_log("📐 Layout opened successfully.", "success")
        except Exception as e:
            self.append_log(f"❌ Error opening layout or setting atlas layer: {str(e)}", "error")


    def open_9_2_layout(self, dest_qpt, parcel_layer):
        """Called in GUI thread — SAFE to use iface here."""
        try:

            project = QgsProject.instance()
            layout_mgr = project.layoutManager()

            new_layout  = QgsPrintLayout(project)
            new_layout .initializeDefaults()

            doc = QDomDocument()
            with open(dest_qpt, "r", encoding="utf-8") as f:
                doc.setContent(f.read())

            context = QgsReadWriteContext()
            new_layout.loadFromTemplate(doc, context)

            new_layout.setName("PPM_With_9(2)_Notice") # IMPORTANT — set layout name AFTER loading template
            layout_mgr.addLayout(new_layout)
            
            # Force provider + renderer + labeling refresh
            parcel_layer.triggerRepaint()

            # supproting_file.triggerRepaint()

            # Set the atlas coverage layer if layout has atlas
            atlas = new_layout.atlas()
            if atlas.enabled():
                atlas.setCoverageLayer(parcel_layer)
                self.append_log(f"🗺️ Atlas coverage layer set to: {parcel_layer.name()}", "success")

            # Open layout designer
            self.iface.openLayoutDesigner(new_layout)

            self.append_log("📐 Layout opened successfully.", "success")
        except Exception as e:
            self.append_log(f"❌ Error opening layout or setting atlas layer: {str(e)}", "error")

                
    def open_vm_layout(self, dest_qpt, parcel_layer, supproting_file):
        """Called in GUI thread — SAFE to use iface here."""
        try:

            project = QgsProject.instance()
            layout_mgr = project.layoutManager()

            new_layout  = QgsPrintLayout(project)
            new_layout .initializeDefaults()

            doc = QDomDocument()
            with open(dest_qpt, "r", encoding="utf-8") as f:
                doc.setContent(f.read())

            context = QgsReadWriteContext()
            new_layout.loadFromTemplate(doc, context)

            new_layout.setName("PPM_VM_A0_TEMPLATE_NEW") # IMPORTANT — set layout name AFTER loading template
            layout_mgr.addLayout(new_layout)
            
            # Force provider + renderer + labeling refresh
            parcel_layer.triggerRepaint()

            supproting_file.triggerRepaint()

            # self.iface.mapCanvas().refreshAllLayers()
            # QgsApplication.processEvents()
            # self.iface.mapCanvas().refresh()


            # Set the atlas coverage layer if layout has atlas
            atlas = new_layout.atlas()
            if atlas.enabled():
                atlas.setCoverageLayer(supproting_file)
                self.append_log(f"🗺️ Atlas coverage layer set to: {supproting_file.name()}", "success")

            # Open layout designer
            self.iface.openLayoutDesigner(new_layout)

            self.append_log("📐 Layout opened successfully.", "success")
            
        except Exception as e:
            self.append_log(f"❌ Error opening layout or setting atlas layer: {str(e)}", "error")
              
    def run_process(self):
        try:        
            # Switch to LOG tab
            self.buttonBox_2.setEnabled(False)
            self.tabWidget.setCurrentIndex(1)
            self.te_log.clear()
            self.progress_bar.setValue(0)
            self.te_log.append("\n\n====================================================\n\n"
                            "\n======        🔹 Starting Maps Generation Steps        ======\n"
                            "\n\n====================================================\n\n")

            errors = []

            # -------------------------------
            # 1. DISTRICT
            # -------------------------------
            district = self.cb_district.currentText().strip()
            if district == "" or district == "-- Select District --":
                errors.append("District not selected.")

            # -------------------------------
            # 2. MANDAL
            # -------------------------------
            mandal = self.le_mandal.text().strip()
            if mandal == "":
                errors.append("Mandal name is required.")

            # -------------------------------
            # 3. VILLAGE
            # -------------------------------
            village = self.le_revenue_village.text().strip()
            if village == "":
                errors.append("Revenue village is required.")

            # -------------------------------
            # 4. LGD CODE
            # -------------------------------
            lgd = self.le_revenue_lgd_code.text().strip()
            if lgd == "":
                errors.append("Revenue LGD code is required.")

            # -------------------------------
            # 5. PARCEL SHAPEFILE
            # -------------------------------
            parcel_layer = self.cb_parcel_shp.currentText().strip()
            if parcel_layer == "" or parcel_layer == "-- Select Layer --":
                errors.append("Parcel shapefile not selected.")

            # -------------------------------
            # 6. BUILTUP SHAPEFILE
            # -------------------------------
            builtup_layer = self.cb_builtup_shp.currentText().strip()
            if builtup_layer == "" or builtup_layer == "-- Select Layer --":
                errors.append("Built-up shapefile not selected.")

            # -------------------------------
            # 6.1 EXTRA LAYER VALIDATIONS
            # -------------------------------

            parcel_qgs_layer = self.get_layer_by_name(parcel_layer)
            builtup_qgs_layer = self.get_layer_by_name(builtup_layer)

            # Ensure layers exist
            if parcel_qgs_layer is None:
                errors.append(f"Parcel layer '{parcel_layer}' not found in project.")

            if builtup_qgs_layer is None:
                errors.append(f"Built-up layer '{builtup_layer}' not found in project.")

            # Only run validations if both loaded
            if parcel_qgs_layer and builtup_qgs_layer:

                # ---------------------------------------
                # ✔ Validation 1: Same file path
                # ---------------------------------------
                parcel_path = parcel_qgs_layer.source().split("|")[0]
                builtup_path = builtup_qgs_layer.source().split("|")[0]

                if parcel_path == builtup_path:
                    errors.append("Parcel and Built-up shapefiles refer to the SAME file path. Please select different files.")

                # ---------------------------------------
                # ✔ Validation 2: Parcel must be polygon
                # ---------------------------------------
                if parcel_qgs_layer.geometryType() != QgsWkbTypes.PolygonGeometry:
                    errors.append(f"Parcel layer '{parcel_layer}' is not a polygon layer.")

                # ---------------------------------------
                # ✔ Validation 3: Builtup must be polygon
                # ---------------------------------------
                if builtup_qgs_layer.geometryType() != QgsWkbTypes.PolygonGeometry:
                    errors.append(f"Built-up layer '{builtup_layer}' is not a polygon layer.")

                # ---------------------------------------
                # ✔ Validation 4: CRS mismatch warning
                # ---------------------------------------
                if parcel_qgs_layer.crs().authid() != builtup_qgs_layer.crs().authid():
                    errors.append(
                        f"CRS mismatch detected:\n"
                        f"Parcel Layer CRS = {parcel_qgs_layer.crs().authid()}\n"
                        f"Built-up Layer CRS = {builtup_qgs_layer.crs().authid()}\n\n"
                        f"Please reproject one of the layers so both use the same CRS."
                    )

                # ---------------------------------------
                # ✔ Validation 5: Project CRS must match both layers
                # ---------------------------------------
                project_crs = QgsProject.instance().crs()

                if parcel_qgs_layer.crs().authid() != project_crs.authid():
                    errors.append(
                        f"Project CRS mismatch:\n"
                        f"Project CRS = {project_crs.authid()}\n"
                        f"Parcel Layer CRS = {parcel_qgs_layer.crs().authid()}\n\n"
                        f"Please set the project CRS to match the layer CRS."
                    )

                if builtup_qgs_layer.crs().authid() != project_crs.authid():
                    errors.append(
                        f"Project CRS mismatch:\n"
                        f"Project CRS = {project_crs.authid()}\n"
                        f"Built-up Layer CRS = {builtup_qgs_layer.crs().authid()}\n\n"
                        f"Please set the project CRS to match the layer CRS."
                    )

            # -------------------------------
            # 7. INITIAL PPN COLUMN
            # -------------------------------
            ppn_column = self.cb_initial_ppn.currentText().strip()
            if ppn_column == "" or ppn_column == "-- Select Column --":
                errors.append("Initial PPN / Chalta No. column not selected.")

            # --------------------------------------------------------
            # 8. EXCEL & Builtup reference column MANDATORY FOR 9(2)
            # --------------------------------------------------------
            excel_path = self.le_excel_path.text().strip()
            excel_required = self.chk_ppms_92_notice.isChecked()
            if excel_required and excel_path == "":
                errors.append("Additional attributes Excel is required because PPMs with 9(2) Notice is enabled.")

            builtup_ppn_column = self.cb_builtup_ppn.currentText().strip()
            if excel_required :
                if builtup_ppn_column == "" or builtup_ppn_column == "-- Select Column --":
                    errors.append("Builtup PPN refrence column not selected.")
                
            # -------------------------------
            # 9. OUTPUT PATH
            # -------------------------------
            output_folder = self.le_output_path.text().strip()
            if output_folder == "":
                errors.append("Output folder path is required.")

            # -------------------------------
            # 10. AT LEAST ONE PROCESS OPTION REQUIRED
            # -------------------------------
            initial_ppm_enabled = self.chk_initial_ppms.isChecked()
            ppms_92_enabled = self.chk_ppms_92_notice.isChecked()
            village_map_enabled = self.chk_village_map.isChecked()

            if not (initial_ppm_enabled or ppms_92_enabled or village_map_enabled):
                errors.append("Please select at least one option: Initial PPM, PPMs with 9(2) Notice, or Village Map.")

            # -----------------------------------------
            # If errors exist → show error + stop
            # -----------------------------------------
            if errors:
                self.append_log("❌ Validation Failed:\n", "error")

                for err in errors:
                    self.append_log(f"❌ {err}", "error")

                # Error popup
                msg = QtWidgets.QMessageBox()
                msg.setIcon(QtWidgets.QMessageBox.Warning)
                msg.setWindowTitle("Validation Error")
                msg.setText("Please fix the following errors before running:")
                msg.setDetailedText("\n".join(errors))
                msg.exec_()

                # Reset PPN dropdown if column issue found
                if "Initial PPN / Chalta No. column not selected." in errors:
                    index = self.cb_initial_ppn.findText("-- Select Column --")
                    if index >= 0:
                        self.cb_initial_ppn.setCurrentIndex(index)

                return  # STOP HERE

            # -------------------------------
            # Collect parameters
            # -------------------------------
            params = {
                "district": district,
                "mandal": mandal,
                "village": village,
                "lgd": lgd,
                "parcel_layer": parcel_layer,
                "builtup_layer": builtup_layer,
                "ppn_column": ppn_column,
                "excel_path": excel_path,
                "excel_required": excel_required,
                "builtup_ppn_column": builtup_ppn_column,
                "output_folder": output_folder,
                "initial_ppm": initial_ppm_enabled,
                "ppms_92_notice": ppms_92_enabled,
                "village_map": village_map_enabled
            }
            p = params
            
            # -----------------------------------------------
            # Determine what to run
            # -----------------------------------------------
            PROCESS_STEPS = {
                "initial_steps": 7,
                "initial_ppm": 17,
                "ppms_92_notice": 18,
                "village_map": 7
            }

            selected_processes = ["initial_steps"]  # base workflow

            if p["initial_ppm"]:
                selected_processes.append("initial_ppm")
            if p["ppms_92_notice"]:
                selected_processes.append("ppms_92_notice")
            if p["village_map"]:
                selected_processes.append("village_map")

            self.total_steps = sum(PROCESS_STEPS[x] for x in selected_processes)
            self.current_step = 0

            self.append_log(f"➡ Running {len(selected_processes)} processing blocks", "info")
            self.append_log(f"🔢 Total steps to run: {self.total_steps}", "info")
            
            # -------------------------------
            # If all OK → Continue process
            # -------------------------------
            self.advance_progress("Initial step 1 👉🏻 ✅ Validation successful. Starting map generation...", "success")

            self.append_log("Starting Maps Generation Process...", "info")

            # ------------------------------------------
            # STARTING BLOCK (combined initial 5 steps)
            # ------------------------------------------
            self.append_log(f"District : {p['district']}", "info")
            self.append_log(f"Mandal : {p['mandal']}", "info")
            self.append_log(f"Village : {p['village']}", "info")
            self.append_log(f"Revenue L.G.D Code : {p['lgd']}", "info")
            self.append_log(f"Scanning inputs → Parcel: {p['parcel_layer']}, Builtup: {p['builtup_layer']}", "info")
            self.append_log(f"Selected PPN column : {p['ppn_column']}", "info")
            if not excel_path:
                pass
            else:
                self.append_log(f"Selected excel path : {p['excel_path']}", "info")
            self.append_log(f"Project Saved Path : {p['output_folder']}", "info")
            self.append_log("Validating input parameters...", "info")
            self.append_log("Initializing workspace...", "info")
            self.append_log("Inputs ready for processing.", "success")




            # ------------------------------------------
            # CREATE FOLDER STRUCTURE
            # ------------------------------------------
            base_dir = p['output_folder']
            district = p['district']
            mandal = p['mandal']
            village = p['village']
            lgd = p['lgd']

            self.folder_lv1 = os.path.join(base_dir, district)
            self.folder_lv2 = os.path.join(self.folder_lv1, mandal)
            self.folder_lv3 = os.path.join(self.folder_lv2, f"{lgd}_{village}")
            self.folder_lv4 = os.path.join(self.folder_lv3, '9(2) Notices')
            self.folder_support = os.path.join(self.folder_lv3, "Supporting Files")
            self.folder_shp = os.path.join(self.folder_support, "Shapefiles")

            try:
                for fd in [self.folder_lv1, self.folder_lv2, self.folder_lv3, self.folder_support, self.folder_shp]:
                    os.makedirs(fd, exist_ok=True)
                self.append_log(f"📁 Folder structure created at: {self.folder_lv3}", "success")
            except Exception as e:
                self.append_log(f"❌ Error creating project folders.")
                raise
            # ------------------------------------------
            # SAVE PROJECT FILE
            # ------------------------------------------
            self.project_path = os.path.join(self.folder_lv3, f"{lgd}_{village}_PPMs.qgz")
            QgsProject.instance().write(self.project_path)

            self.advance_progress(f"Initial step 2 👉🏻 ✅ 💾 Project saved: {self.project_path}", "success")

            # ---------------------------------------------------
            #  EXCEL HANDLING (if uploaded)
            # ---------------------------------------------------
            # excel_path = self.le_excel_path.text().strip()
            # excel_required = self.chk_ppms_92_notice.isChecked()

            if excel_required:

                # Must be provided
                if excel_path == "":
                    errors.append("Additional attributes Excel is required because PPMs with 9(2) Notice is enabled.")
                else:
                    try:
                        # 🔹 Save Excel inside: self.folder_lv4 (9(2) Notices)
                        os.makedirs(self.folder_lv4, exist_ok=True)

                        excel_filename = os.path.basename(excel_path)
                        saved_excel_path = os.path.join(self.folder_lv4, excel_filename)

                        # Copy Excel to project folder
                        shutil.copyfile(excel_path, saved_excel_path)
                        self.append_log(f"📄 Excel saved at: {saved_excel_path}", "success")

                        # --------------------------------------
                        # LOAD & VALIDATE REQUIRED COLUMNS
                        # --------------------------------------
                        self.append_log("🔎 Validating Excel contents...", "info")

                        import pandas as pd
                        try:
                            df = pd.read_excel(saved_excel_path)
                        except Exception as err:
                            raise Exception(f"Could not read Excel file: {err}")

                        required_cols = [
                            'PPN',
                            'Property type (Individual/Joint/Apartment/Government)',
                            'Panchayat Name',
                            'Owner Name',
                            'Relation (W/O,H/O,S/O,D/O)',
                            'Assessment No.',
                            'Remarks'
                        ]

                        missing = [c for c in required_cols if c not in df.columns]

                        if len(missing) > 0:
                            error_text = (
                                "❌ Excel is missing required column(s):\n"
                                + "\n".join([f"• {m}" for m in missing])
                            )
                            self.append_log(error_text, "error")
                            raise Exception("Excel validation failed due to missing headers.")

                        self.append_log("✅ Excel headers verified successfully.", "success")

                        # --------------------------------------
                        # LOAD EXCEL AS A QGIS TABLE LAYER
                        # --------------------------------------
                        self.append_log("📥 Loading Excel into QGIS...", "info")

                        excel_norm = os.path.normpath(saved_excel_path).replace("\\", "/")

                        try:
                            # Use OGR-style sheet loading (most reliable)
                            excel_uri = f"{excel_norm}|layername=Sheet1"
                            self.excel_layer = QgsVectorLayer(excel_uri, "PPM_92_Excel", "ogr")

                            if not self.excel_layer.isValid():
                                raise Exception(
                                    "Excel could not be loaded. Verify:\n"
                                    "• Sheet name is 'Sheet1'\n"
                                    "• File is .xlsx format\n"
                                    "• File is not password protected"
                                )

                            QgsProject.instance().addMapLayer(self.excel_layer)
                            self.append_log("✅ Excel loaded successfully into QGIS.", "success")

                        except Exception as e:
                            self.append_log(f"❌ Excel processing error: {e}", "error")
                            raise
                    except Exception as e:
                        self.append_log(f"❌ Excel processing error: {e}", "error")
                        raise


            # ------------------------------------------
            # EXPORT SHAPEFILES
            # ------------------------------------------
                
            self.parcel_layer = QgsProject.instance().mapLayersByName(p['parcel_layer'])[0]
            self.builtup_layer = QgsProject.instance().mapLayersByName(p['builtup_layer'])[0]

            self.parcel_out = os.path.join(self.folder_shp, f"{lgd}_Parcel_Area.gpkg")
            self.builtup_out = os.path.join(self.folder_shp, f"{lgd}_Builtup_Area.gpkg")

            # Export both layers safely
            self.parcel_new = self.safe_export_layer(self.parcel_layer, self.parcel_out, f"{lgd}_Parcel Area")
            self.builtup_new = self.safe_export_layer(self.builtup_layer, self.builtup_out, f"{lgd}_Builtup Area")

            self.advance_progress(f"Initial step 3 👉🏻 ✅ 💾 Parcel and Built-up shape files stored and added back to project", "success")

            if not self.parcel_new or not self.builtup_new:
                self.append_log("❌ Export failed. Stopping process.", "error")
                raise Exception ("Maps Generation process stopped due to export failure")

            # ------------------------------------------
            # LINK PARCEL & BUILTUP USING USER SELECTED PPN COLUMN
            # ------------------------------------------

            self.append_log("🔗 Creating Parcel Reference Column...", "info")

            try:
                # ------------------------------
                # Step 1: Create new fields
                # ------------------------------
                new_fields = {
                    "PAR_REF": QVariant.Int,
                    "DISTRICT": QVariant.String,
                    "MANDAL": QVariant.String,
                    "VILLAGE": QVariant.String,
                    "LGD_CODE": QVariant.String
                }

                if not self.add_fields_safe(self.parcel_new, new_fields):
                    raise Exception("Failed to create fields")

                # ------------------------------
                # Step 2: Prepare attribute updates
                # ------------------------------
                updates = []
                ppn_col_selected = p["ppn_column"]

                for feat in self.parcel_new.getFeatures():
                    updates.append(
                        (feat.id(), {
                            "PAR_REF": feat[ppn_col_selected],
                            "DISTRICT": p["district"],
                            "MANDAL": p["mandal"],
                            "VILLAGE": p["village"],
                            "LGD_CODE": p["lgd"],
                        })
                    )

                # ------------------------------
                # Step 3: Apply the attribute values
                # ------------------------------
                if not self.update_attributes_safe(self.parcel_new, updates):
                    raise Exception("Failed to update attributes")

                # ------------------------------
                # Step 4: Commit the changes
                # ------------------------------
                if not self.commit_layer_safe(self.parcel_new):
                    raise Exception("Failed to commit changes")
                self.advance_progress(f"Initial step 4 👉🏻 ✅ LINK PARCEL & BUILTUP USING USER SELECTED PPN COLUMN", "success")
            except Exception as e:
                self.append_log(f"❌ Error in parcel linking process: {e}", "error")
                return

            # ----------------------------------
            # Area Columns added to Shape file
            # ----------------------------------
            try:
                existing_fields = [f.name() for f in self.parcel_new.fields()]
                provider = self.parcel_new.dataProvider()
                # Add Ar_M2
                if "Ar_M2" not in existing_fields:
                    provider.addAttributes([QgsField("Ar_M2", QVariant.Double, len=20, prec=3)])
                    self.parcel_new.updateFields()
                # Add Ar_Ya2
                if "Ar_Ya2" not in existing_fields:
                    provider.addAttributes([QgsField("Ar_Ya2", QVariant.Double, len=20, prec=3)])
                    self.parcel_new.updateFields()

                # Re-fetch fields
                fields = self.parcel_new.fields()
                area_m_index = fields.indexOf("Ar_M2")
                area_y_index = fields.indexOf("Ar_Ya2")

                # Start editing and calculate geometry-based areas
                if self.parcel_new.isEditable():
                    self.parcel_new.commitChanges()   
                self.parcel_new.startEditing()

                for feature in self.parcel_new.getFeatures():
                    geom = feature.geometry()
                    if geom and not geom.isEmpty():
                        # Calculate area in square meters (QGIS geometry area is always in layer CRS units)
                        area_sq_m = geom.area()
                        
                        # Convert to square yards (1 sqm = 1.19599 sq yards)
                        area_sq_y = area_sq_m * 1.19599

                        # Update attributes
                        self.parcel_new.changeAttributeValue(feature.id(), area_m_index, round(area_sq_m, 2))
                        self.parcel_new.changeAttributeValue(feature.id(), area_y_index, round(area_sq_y, 2))

                self.parcel_new.commitChanges()
                self.parcel_new.triggerRepaint()
                QgsApplication.processEvents()
                
                existing_fields_b = [f.name() for f in self.builtup_new.fields()]
                provider_b = self.builtup_new.dataProvider()
                # Add Ar_M2
                if "Ar_M2" not in existing_fields_b:
                    provider_b.addAttributes([QgsField("Ar_M2", QVariant.Double, len=20, prec=3)])
                    self.builtup_new.updateFields()
                # Add Ar_Ya2
                if "Ar_Ya2" not in existing_fields_b:
                    provider_b.addAttributes([QgsField("Ar_Ya2", QVariant.Double, len=20, prec=3)])
                    self.builtup_new.updateFields()

                # Re-fetch fields
                fields_b = self.builtup_new.fields()
                area_m_index_b = fields_b.indexOf("Ar_M2")
                area_y_index_b = fields_b.indexOf("Ar_Ya2")

                # Start editing and calculate geometry-based areas
                if self.builtup_new.isEditable():
                    self.builtup_new.commitChanges()   
                self.builtup_new.startEditing()

                for feature in self.builtup_new.getFeatures():
                    geom = feature.geometry()
                    if geom and not geom.isEmpty():
                        # Calculate area in square meters (QGIS geometry area is always in layer CRS units)
                        area_sq_m = geom.area()
                        
                        # Convert to square yards (1 sqm = 1.19599 sq yards)
                        area_sq_y = area_sq_m * 1.19599

                        # Update attributes
                        self.builtup_new.changeAttributeValue(feature.id(), area_m_index_b, round(area_sq_m, 2))
                        self.builtup_new.changeAttributeValue(feature.id(), area_y_index_b, round(area_sq_y, 2))

                self.builtup_new.commitChanges()
                self.builtup_new.triggerRepaint()
                QgsApplication.processEvents()               
                
                
                self.advance_progress(f"Initial step 5 👉🏻 ✅ Area Columns added to Shape file", "success")
            except Exception as e:
                self.append_log(f"❌ Error in creating Area Fields: {e}", "error")
                return



            # ------------------------------------------
            # SPATIAL LINKING — INTERSECT BUILTUP WITH PARCELS
            # ------------------------------------------
            try:
                self.append_log("🔄 Linking Built-up with Parcel Reference using Intersection...", "info")

                # Paths (use GPKG to avoid shapefile limitations)
                self.parcel_ref_path = os.path.join(self.folder_shp, f"{lgd}_ParcelArea_withRef.gpkg")
                self.builtup_in_path = os.path.join(self.folder_shp, f"{lgd}_Builtup_Input.gpkg")
                self.builtup_ref_path_link = os.path.join(self.folder_shp, f"Builtup_Reference_link.gpkg")
                self.builtup_ref_path_out = os.path.join(self.folder_shp, f"Builtup_Reference.gpkg")
                reference_col = "PAR_REF"

                # Ensure any pending edits are committed before exporting
                if self.parcel_new.isEditable():
                    if not self.parcel_new.commitChanges():
                        self.parcel_new.rollBack()
                        raise Exception("Failed to commit parcel_new before exporting.")
                if self.builtup_new.isEditable():
                    if not self.builtup_new.commitChanges():
                        self.builtup_new.rollBack()
                        raise Exception("Failed to commit builtup_new before exporting.")

                QgsApplication.processEvents()

                # Export parcel_new to GPKG on disk
                err, msg = QgsVectorFileWriter.writeAsVectorFormat(
                    self.parcel_new,
                    self.parcel_ref_path,
                    "UTF-8",
                    self.parcel_new.crs(),
                    "GPKG"
                )
                if err != QgsVectorFileWriter.NoError:
                    raise Exception(f"Failed to export parcel layer to disk: {msg}")

                # Export builtup_new to GPKG on disk (use as INPUT for intersection)
                err2, msg2 = QgsVectorFileWriter.writeAsVectorFormat(
                    self.builtup_new,
                    self.builtup_in_path,
                    "UTF-8",
                    self.builtup_new.crs(),
                    "GPKG"
                )
                if err2 != QgsVectorFileWriter.NoError:
                    raise Exception(f"Failed to export builtup layer to disk: {msg2}")

                QgsApplication.processEvents()

                # Reload the on-disk layers (fresh instances) and validate
                self.parcel_ref_layer = QgsVectorLayer(self.parcel_ref_path, f"{lgd}_ParcelArea_withRef", "ogr")
                self.builtup_in_layer = QgsVectorLayer(self.builtup_in_path, f"{lgd}_Builtup_Input", "ogr")

                if not self.parcel_ref_layer.isValid():
                    raise Exception("Reloaded parcel_ref_layer is invalid after export.")
                if not self.builtup_in_layer.isValid():
                    raise Exception("Reloaded builtup_in_layer is invalid after export.")

                # Use file paths or datasource URIs for processing (safer)
                # Many QGIS setups accept the string path directly. Use layer.source() as fallback.
                input_source = self.builtup_in_path
                overlay_source = self.parcel_ref_path

                # Run intersection with on-disk inputs and on-disk output
                processing_result = processing.run(
                    "native:intersection",
                    {
                        "INPUT": input_source,
                        "OVERLAY": overlay_source,
                        "INPUT_FIELDS": [],          # fields from INPUT to keep (empty means none)
                        "OVERLAY_FIELDS": [reference_col],  # keep parcel ref from overlay
                        "OUTPUT": self.builtup_ref_path_link
                    }
                )

                # Check processing output
                if "OUTPUT" not in processing_result or not os.path.exists(self.builtup_ref_path_link):
                    raise Exception("Intersection processing failed or output missing.")

                self.append_log("✔ Built-up polygons successfully linked to Parcel Reference.", "success")

                # Reload the intersection result to a proper layer object and export to final output if required
                intersect_layer = QgsVectorLayer(self.builtup_ref_path_link, "Builtup_Reference_tmp", "ogr")
                if not intersect_layer.isValid():
                    raise Exception("Intersection output layer is invalid.")

                # Optionally export (or simply move) the result to a stable final file using safe_export_layer
                self.builtup_ref_layer = self.safe_export_layer(intersect_layer, self.builtup_ref_path_out, "Builtup_Reference")
                # -------------------------------------------------------------
                # CLEAN UP TEMPORARY FILES
                # -------------------------------------------------------------
                try:
                    self.append_log("🧹 Cleaning up temporary processing files...", "info")

                    temp_files = [
                        self.builtup_in_path,
                        self.builtup_ref_path_link
                    ]

                    for f in temp_files:
                        if os.path.exists(f):
                            try:
                                os.remove(f)
                                self.append_log(f"   • Removed temp file: {os.path.basename(f)}", "success")
                            except Exception as remove_err:
                                # Soft warning only; plugin should NOT stop
                                # self.append_log(f"⚠ Could not delete temp file: {f} ({remove_err})", "info")
                                pass

                    self.append_log("✔ Temporary files cleaned successfully.", "success")

                except Exception as cleanup_err:
                    # Never stop the plugin here
                    self.append_log(f"⚠ Warning: Cleanup issue (not critical): {cleanup_err}", "warning")


                if not self.builtup_ref_layer:
                    self.append_log("❌ Builtup_Reference.gpkg export failed. Stopping process.", "error")
                    raise Exception("Maps Generation process stopped due to export Builtup_Reference failure")

                self.advance_progress(f"Initial step 6 👉🏻 ✅ SPATIAL LINKING — INTERSECT BUILTUP WITH PARCELS", "success")
            except Exception as e:
                self.append_log(f"❌ Error in creating spatial linking BUILTUP with PARCELS: {e}", "error")
                return

            # -------------------------------------------------------------
            # REMOVE TINY PLINTH POLYGONS + RECREATE SPATIAL INDEX
            # -------------------------------------------------------------
            self.append_log("🧹 Removing tiny built-up polygons (cleaning)...", "info")
            # Load actual layer
            # builtup_ref_layer = QgsVectorLayer(builtup_ref_path, "Builtup_Reference", "ogr")

            try:
                # builtup_ref_layer = QgsVectorLayer(builtup_ref_path, "Builtup_Reference", "ogr")
                if not self.builtup_ref_layer.isValid():
                    raise Exception("Built-up linked shapefile could not be loaded for cleaning.")

                # Open edit session
                if self.builtup_ref_layer.isEditable():
                    self.builtup_ref_layer.commitChanges()
                self.builtup_ref_layer.startEditing()

                remove_count = 0
                min_area = 2.0  # sq.m — adjust if needed

                for f in self.builtup_ref_layer.getFeatures():
                    if f.geometry().area() < min_area:
                        self.builtup_ref_layer.deleteFeature(f.id())
                        remove_count += 1

                # Commit edits
                ok = self.builtup_ref_layer.commitChanges()
                if not ok:
                    self.builtup_ref_layer.rollBack()
                    raise Exception("Failed to commit deletions of tiny polygons.")

                self.append_log(f"✔ Removed {remove_count} tiny built-up polygons (< {min_area} sqm)", "success")

            except Exception as e:
                self.append_log(f"❌ Error while removing small plinth polygons: {str(e)}", "error")
                raise
            
            # -------------------------------------------------------------
            # RECREATE SPATIAL INDEX FOR BUILT-UP REFERENCE LAYER
            # -------------------------------------------------------------
            self.append_log("📐 Rebuilding spatial index for cleaned built-up layer...", "info")

            try:
                # Reload layer (fresh after deletions)
                # builtup_ref_layer = QgsVectorLayer(builtup_ref_path, "Builtup_Reference", "ogr")

                if not self.builtup_ref_layer.isValid():
                    raise Exception("Builtup_Reference layer invalid after cleaning.")

                index = QgsSpatialIndex(self.builtup_ref_layer.getFeatures())
                self.append_log("✔ Spatial index successfully recreated.", "success")
                self.append_log("✔ Built-up polygons successfully linked to Parcel Reference removing tiny features.", "success")
                self.advance_progress(f"Initial step 7 👉🏻 ✅ REMOVE TINY PLINTH POLYGONS + RECREATE SPATIAL INDEX", "success")
        
            except Exception as e:
                self.append_log(f"❌ Spatial index creation failed: {str(e)}", "error")
                raise

            # ---------------------------------------------------------
            # Initial Steps Are Completed
            # ---------------------------------------------------------

            self.append_log("✅ All initial steps completed successfully.", "success")

            # # Update progress bar to reflect completion of initial steps
            # initial_steps_count = PROCESS_STEPS["initial_steps"]

            # progress_percent = int((initial_steps_count / self.total_steps) * 100)
            # self.progress_bar.setValue(progress_percent)
            # QgsApplication.processEvents()


            # -----------------------------------------------
            # Run Initial PPM
            # -----------------------------------------------
            if "initial_ppm" in selected_processes:
                self.run_initial_ppm(p)

            # -----------------------------------------------
            # Run 92 Notice
            # -----------------------------------------------
            if "ppms_92_notice" in selected_processes:
                self.run_ppms_92(p)

            # -----------------------------------------------
            # Run Village Map
            # -----------------------------------------------
            if "village_map" in selected_processes:
                self.run_village_map(p)

            # ------------------------------------------
            # FINISH
            # ------------------------------------------

            self.append_log(" ", "info")
            self.append_log(" ", "info")
            self.append_log("====================================================", "info")
            self.append_log(" ", "info")
            self.append_log("=========  🎯 All Requests completed successfully.  =========", "info")
            self.append_log("", "info")
            self.append_log("====================================================", "info")
            self.append_log(" ", "info")
            self.append_log(" ", "info")
            self.buttonBox_2.setEnabled(True)



        except Exception as e:
            self.append_log(f"❌ Error in processing user selected process : {e}", "error")
            self.buttonBox_2.setEnabled(True)

    def advance_progress(self, msg="", msg_type="info"):
        """
        Increment progress bar and optionally log a message.

        :param msg: Message to log
        :param msg_type: Type of message ('info', 'success', 'error', etc.)
        """
        self.current_step += 1
        percent = int((self.current_step / self.total_steps) * 100)
        self.progress_bar.setValue(percent)
        QgsApplication.processEvents()
        
        if msg:
            self.append_log(msg, msg_type)

    def normalize_ppn(self,val):
        if val is None:
            return None
        try:
            return str(int(float(val)))
        except Exception:
            return str(val).strip()

    def run_initial_ppm(self,p):
        self.append_log(" ", "info")
        self.append_log(" ", "info")
        self.append_log("====================================================", "info")
        self.append_log(" ", "info")
        self.append_log("========  🗺️ Initial PPM generation process starts  ========", "info")
        self.append_log("", "info")
        self.append_log("====================================================", "info")
        self.append_log(" ", "info")
        self.append_log(" ", "info")
        
        # ------------------------------------------
        # 1) Boundary from parcel shapefile & builtup shapefile
        # Output: Parcel_Boundary.shp & Builtup_Boundary.shp in folder_shp
        # ------------------------------------------
        try:
            parcel_boundary_path = os.path.join(self.folder_shp, "Parcel_Boundary.shp")
            params_parcel_boundary = {
                'INPUT': self.parcel_new,
                'OUTPUT': parcel_boundary_path
            }
            self.append_log("⌛ Running Parcel Boundary algorithm...", "info")
            res_boundary = processing.run("native:boundary", params_parcel_boundary)
            self.advance_progress(f"📤 Parcel Boundary created: {parcel_boundary_path}", "success")

            builtup_boundary_path = os.path.join(self.folder_shp, "Builtup_Boundary.shp")
            params_builtup_boundary = {
                'INPUT': self.builtup_ref_layer,
                'OUTPUT': builtup_boundary_path
            }
            self.append_log("⌛ Running Built-up Boundary algorithm...", "info")
            res_boundary = processing.run("native:boundary", params_builtup_boundary)
            self.advance_progress(f"📤 Built-up Boundary created: {builtup_boundary_path}", "success")
        except Exception as e:
            self.append_log(f"❌ Error creating Parcel Boundary & Built-up Boundary: {e}", "error")

        # ------------------------------------------
        # 2) Explode lines 
        # ------------------------------------------
        try:
            # ------------------------------------------
            # 2A) Explode lines (Parcel_Boundary -> Parcel_Explode_Lines.shp)
            # ------------------------------------------
            parcel_explode = os.path.join(self.folder_shp, "Parcel_Explode_Lines.shp")
            params_explode1 = {
                'INPUT': parcel_boundary_path,
                'OUTPUT': parcel_explode
            }
            self.append_log("⌛ Exploding Parcel boundary to lines...", "info")
            res_explode1 = processing.run("native:explodelines", params_explode1)
            self.advance_progress(f"📤 Exploded Parcel lines exported: {parcel_explode}", "success")
            parcel_explode_layer_name = "Parcel_Explode_Lines"
            parcel_explode_path = self.readding_layer(parcel_explode_layer_name, parcel_explode)

            # ------------------------------------------
            # 2B) Explode lines (Builtup_Boundary -> Builtup_Explode_Lines.shp)
            # ------------------------------------------
            builtup_explode = os.path.join(self.folder_shp, "Builtup_Explode_Lines.shp")
            params_explode = {
                'INPUT': builtup_boundary_path,
                'OUTPUT': builtup_explode
            }
            self.append_log("⌛ Exploding Builtup boundary to lines...", "info")
            res_explode1 = processing.run("native:explodelines", params_explode)
            self.advance_progress(f"📤 Exploded Builtup lines exported: {builtup_explode}", "success")
            builtup_explode_layer_name = "Builtup_Explode_Lines"
            builtup_explode_path = self.readding_layer(builtup_explode_layer_name, builtup_explode)

        except Exception as e:
            self.append_log(f"❌ Error creating Explode lines for Parcel Boundary & Built-up Boundary: {e}", "error")

        try:
            # ------------------------------------------
            # 3) Extract vertices from Parcel_Boundary -> Parcel_Vertices.shp
            # ------------------------------------------
            parcel_vertices_path = os.path.join(self.folder_shp, "Parcel_Vertices.shp")
            params_vertices = {
                'INPUT': parcel_boundary_path,
                'OUTPUT': parcel_vertices_path
            }
            self.append_log("⌛ Extracting vertices...", "info")
            res_vertices = processing.run("native:extractvertices", params_vertices)
            self.advance_progress(f"📤 Vertices extracted: {parcel_vertices_path}", "success")
        except Exception as e:
            self.append_log(f"❌ Error creating vertices for Parcel layer: {e}", "error")


        # ------------------------------------------
        # 4) Multi-ring buffer around vertices
        #    (NUMBER=1, DISTANCE=2, SEGMENTS=8) -> Parcel_Buffer.shp
        # Note: native:multiringbuffer expects a comma-separated DISTANCES string
        # ------------------------------------------

        parcel_buffer_path = os.path.join(self.folder_shp, "Parcel_Buffer.shp")

        self.append_log("⌛ Creating multi-ring buffer around vertices...", "info")
        try:
            # Try native:multiringbuffer first
            params_buffer = {
                'INPUT': parcel_vertices_path,
                'DISTANCE': 2,   # distance per ring
                'NUMBER': 1,     # number of rings
                'SEGMENTS': 8,
                'OUTPUT': parcel_buffer_path
            }
            res_buffer = processing.run("native:multiringbuffer", params_buffer)
            self.advance_progress(f"📤 Multi-ring buffer created: {parcel_buffer_path}", "success")
        except Exception as e1:
            # self.append_log(f"⚠️ native:multiringbuffer failed ({str(e1)}), trying native:buffer...", "error")
            try:
                # 2️⃣ Fallback to simple buffer
                params_fallback = {
                    'INPUT': parcel_vertices_path,
                    'DISTANCE': 1,
                    'SEGMENTS': 8,
                    'DISSOLVE': False,
                    'OUTPUT': parcel_buffer_path
                }
                res = processing.run("native:buffer", params_fallback)
                self.advance_progress(f"📤 Buffer created using fallback: {parcel_buffer_path}", "success")

            except Exception as e2:
                # self.append_log(f"⚠️ native:buffer also failed ({str(e2)}), trying native:multiringconstantbuffer...", "error")
                # 3️⃣ Last fallback: native:multiringconstantbuffer
                params_constant = {
                    'INPUT': parcel_vertices_path,
                    'RINGS': 1,
                    'DISTANCE': 1,
                    'OUTPUT': parcel_buffer_path
                }
                res = processing.run("native:multiringconstantbuffer", params_constant)
                self.advance_progress(f"📤 Multi-ring constant buffer created: {parcel_buffer_path}", "success")

        # ------------------------------------------
        # 5) Clip: parcel_explode (input) clipped by parcel_buffer (overlay) -> Parcel_Clip.shp
        # ------------------------------------------
        try:
            parcel_clip_path = os.path.join(self.folder_shp, "Parcel_Clip.shp")
            params_clip = {
                'INPUT': parcel_explode_path,
                'OVERLAY': parcel_buffer_path,
                'OUTPUT': parcel_clip_path
            }
            self.append_log("⌛ Clipping exploded lines by buffer...", "info")
            res_clip = processing.run("native:clip", params_clip)
            self.advance_progress(f"📤 Clipped layer created: {parcel_clip_path}", "success")
        except Exception as e:
            self.append_log(f"❌ Error creating Clip layer: {e}", "error")

        # ------------------------------------------
        # 6) Explode lines for the clip shapefile -> Parcel_Clip_Explode_Lines.shp
        # ------------------------------------------
        try:
            parcel_clip_explode = os.path.join(self.folder_shp, "Parcel_Clip_Explode_Lines.gpkg")
            params_explode2 = {
                'INPUT': parcel_clip_path,
                'OUTPUT': parcel_clip_explode
            }
            self.append_log("⌛ Exploding clipped lines...", "info")
            res_explode2 = processing.run("native:explodelines", params_explode2)
            self.advance_progress(f"📤 Exploded clipped lines exported: {parcel_clip_explode}", "success")
            Parcel_Clip_Explode_Lines_layer_name = "Parcel_Clip_Explode_Lines"
            parcel_clip_explode_path = self.readding_layer(Parcel_Clip_Explode_Lines_layer_name, parcel_clip_explode)
        except Exception as e:
            self.append_log(f"❌ Error creating Explode lines for the clip shapefile: {e}", "error")

        # ------------------------------------------
        # 7A) Add 'length' field to Parcel_Explode_Lines and populate with geometry length
        # ------------------------------------------
        self.append_log("⌛ Adding 'length' field for Parcel and populating values...", "info")
        
        try:
            # load exploded lines layer (first explode result)
            # Parcel_explode_lines_layer = QgsVectorLayer(parcel_explode_path, "Parcel_Explode_Lines", "ogr")
            Parcel_explode_lines_layer = parcel_explode_path

            if not Parcel_explode_lines_layer.isValid():
                self.append_log("❌ Failed to open Parcel exploded lines layer for length calculation.", "error")
            else:
                dp = Parcel_explode_lines_layer.dataProvider()
                # Add field if possible
                if dp.capabilities() & dp.AddAttributes:
                    Parcel_explode_lines_layer.startEditing()
                    fld = QgsField("length", QVariant.Double, "", 10, 2)
                    dp.addAttributes([fld])
                    Parcel_explode_lines_layer.updateFields()

                    idx = Parcel_explode_lines_layer.fields().indexFromName("length")
                    # iterate and set length
                    for feat in Parcel_explode_lines_layer.getFeatures():
                        geom = feat.geometry()
                        if geom is None:
                            continue
                        length_value = geom.length()
                        Parcel_explode_lines_layer.changeAttributeValue(feat.id(), idx, round(float(length_value), 2))
                    Parcel_explode_lines_layer.commitChanges()
                    self.advance_progress("✅ 'length' field added and populated on Parcel_Explode_Lines.", "success")
                else:
                    self.append_log("❌ Layer provider does not support adding attributes.", "error")
        except Exception as e:
            self.append_log(f"❌ Error Adding 'length' field to parcels : {e}", "error")

        # ------------------------------------------
        # 7B) Add 'length' field to Builtup_Explode_Lines and populate with geometry length
        # ------------------------------------------
        try:
            self.append_log("⌛ Adding 'length' field for Builtup and populating values...", "info")

            # load exploded lines layer (first explode result)
            # Builtup_explode_lines_layer = QgsVectorLayer(builtup_explode_path, "Builtup_Explode_Lines", "ogr")
            Builtup_explode_lines_layer = builtup_explode_path

            if not Builtup_explode_lines_layer.isValid():
                self.append_log("❌ Failed to open Builtup exploded lines layer for length calculation.", "error")
            else:
                dp = Builtup_explode_lines_layer.dataProvider()
                # Add field if possible
                if dp.capabilities() & dp.AddAttributes:
                    Builtup_explode_lines_layer.startEditing()
                    fld = QgsField("length", QVariant.Double, "", 10, 2)
                    dp.addAttributes([fld])
                    Builtup_explode_lines_layer.updateFields()

                    idx = Builtup_explode_lines_layer.fields().indexFromName("length")
                    # iterate and set length
                    for feat in Builtup_explode_lines_layer.getFeatures():
                        geom = feat.geometry()
                        if geom is None:
                            continue
                        length_value = geom.length()
                        Builtup_explode_lines_layer.changeAttributeValue(feat.id(), idx, round(float(length_value), 2))
                    Builtup_explode_lines_layer.commitChanges()
                    self.advance_progress("✅ 'length' field added and populated on Builtup_Explode_Lines.", "success")
                else:
                    self.append_log("❌ Layer provider does not support adding attributes.", "error")
        except Exception as e:
            self.append_log(f"❌ Error Adding 'length' field to built-up : {e}", "error")

        # ============================================================
        # 8) Add fields (point_ID, Easting_X, Northing_Y) to vertices
        # ============================================================
        try:
            self.append_log("⌛ Adding point_ID, Easting_X, Northing_Y fields to vertices...", "info")

            vertices_layer = QgsVectorLayer(parcel_vertices_path, "Parcel_Vertices", "ogr")
            if not vertices_layer.isValid():
                self.append_log("❌ Failed to load Parcel_Vertices for attribute update.", "error")
            else:
                dp = vertices_layer.dataProvider()

                if dp.capabilities() & QgsVectorDataProvider.AddAttributes:
                    vertices_layer.startEditing()
                    fdefs = []

                    if vertices_layer.fields().indexFromName("point_ID") == -1:
                        fdefs.append(QgsField("point_ID", QVariant.Int, '', 10, 0))
                    if vertices_layer.fields().indexFromName("Easting_X") == -1:
                        fdefs.append(QgsField("Easting_X", QVariant.Double, '', 15, 3))
                    if vertices_layer.fields().indexFromName("Northing_Y") == -1:
                        fdefs.append(QgsField("Northing_Y", QVariant.Double, '', 15, 3))

                    if fdefs:
                        dp.addAttributes(fdefs)

                    vertices_layer.updateFields()

                    idx_pid = vertices_layer.fields().indexFromName("point_ID")
                    idx_e = vertices_layer.fields().indexFromName("Easting_X")
                    idx_n = vertices_layer.fields().indexFromName("Northing_Y")

                    has_vertex_ind = vertices_layer.fields().indexFromName("vertex_ind") != -1

                    for i, feat in enumerate(vertices_layer.getFeatures()):
                        geom = feat.geometry()
                        if geom is None:
                            continue

                        pt = geom.asPoint()  # Extract Vertices ALWAYS returns point geometry

                        # point_ID logic
                        if idx_pid != -1:
                            if has_vertex_ind:
                                try:
                                    vertices_layer.changeAttributeValue(feat.id(), idx_pid,
                                                                        int(feat.attribute("vertex_ind")) + 1)
                                except Exception:
                                    vertices_layer.changeAttributeValue(feat.id(), idx_pid, i + 1)
                            else:
                                vertices_layer.changeAttributeValue(feat.id(), idx_pid, i + 1)

                        # Easting/ Northing
                        if pt:
                            if idx_e != -1:
                                vertices_layer.changeAttributeValue(feat.id(), idx_e, round(pt.x(), 3))
                            if idx_n != -1:
                                vertices_layer.changeAttributeValue(feat.id(), idx_n, round(pt.y(), 3))

                    vertices_layer.commitChanges()
                    self.advance_progress("✅ Fields added & populated for Parcel_Vertices.", "success")
                else:
                    self.append_log("❌ Provider does not support attribute additions on vertices.", "error")
        except Exception as e:
            self.append_log(f"❌ Error Adding 'co-ordinates' : {e}", "error")
                        
        # ============================================================
        # 9–11) Remove duplicate vertices (PAR_REF, Easting_X, Northing_Y)
        # Robust version: use geometry coords (not relying on attr columns)
        # ============================================================

        try:
            self.append_log("⌛ Removing duplicate vertices...", "info")

            if not vertices_layer or not vertices_layer.isValid():
                raise Exception("Vertices layer is not loaded or invalid.")

            # --------------------------------------------------------
            # 1) Run duplicate removal
            # --------------------------------------------------------
            no_dup_path_link = os.path.join(self.folder_shp, "Parcel_No_Dup_Vertices.shp")
            if os.path.exists(no_dup_path_link):
                os.remove(no_dup_path_link)

            result = processing.run(
                "native:removeduplicatesbyattribute",
                {
                    'INPUT': vertices_layer,
                    'FIELDS': ['PAR_REF', 'Easting_X', 'Northing_Y'],
                    'OUTPUT': no_dup_path_link
                }
            )
            # --------------------------------------------------------
            # 3) Reload final layer to project
            # --------------------------------------------------------
            no_dup_layer = self.readding_layer("Parcel_No_Dup_Vertices", no_dup_path_link)
            if not no_dup_layer:
                raise Exception("Failed to load the no-duplicate output layer.")

            self.advance_progress("✅ Duplicate vertices removed & saved successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error removing duplicate vertices: {e}", "error")
            raise

        # ============================================================
        # 13) Apply QML Style Files
        # ============================================================
        try:
            self.append_log("🎨 Applying QML styles...", "info")

            style_base = os.path.join(
                os.path.dirname(__file__),
                "styling_properties",
                "Initial Property Parcel Maps"
            )

            # QML mappings
            qml_files = {
                "Parcel_Explode_Lines": "Parcel_explode_lines.qml",
                "Builtup_Explode_Lines": "Builtup_explode_lines.qml",
                "Parcel_Clip_Explode_Lines": "Parcel_clip_explode_lines.qml",
                "Parcel_No_Dup_Vertices": "Parcel_no_dup_vertices.qml",
                "Builtup_Reference": "Builtup_Ref.qml"
            }
            
            added_layers = {
                "Parcel_Explode_Lines": parcel_explode_path,     # actual QgsVectorLayer
                "Builtup_Explode_Lines": builtup_explode_path,
                "Parcel_Clip_Explode_Lines": parcel_clip_explode_path,
                "Parcel_No_Dup_Vertices": no_dup_layer,
                "Builtup_Reference": self.builtup_ref_layer
            }

            
            # Apply styling
            for layer_name, qml_file in qml_files.items():
                if layer_name in added_layers:
                    qml_path = os.path.join(style_base, qml_file)
                    # self.apply_qml_style(added_layers[layer_name], qml_path)
                    self.apply_qml_style(added_layers[layer_name], qml_path)
                else:
                    self.append_log(f"⚠️ Layer not found for styling: {layer_name}", "warning")

            parcel_qml_path = os.path.join(style_base, "Parcel_Polygon.qml")
            self.apply_qml_style(self.parcel_new, parcel_qml_path)

            # Convert file paths inside added_layers to actual QgsVectorLayer objects
            for key, value in added_layers.items():
                if isinstance(value, str):   # file path
                    layer = QgsVectorLayer(value, key, "ogr")
                    if layer and layer.isValid():
                        added_layers[key] = layer
                        QgsProject.instance().addMapLayer(layer)
                    else:
                        self.append_log(f"⚠️ Failed to load layer from: {value}", "warning")

            self.advance_progress(f"✅ QML Style Applied to all required fiels", "success")

        except Exception as e:
            self.append_log(f"❌ Error Applying QML styles", "error")
            raise
        
        # ============================================================
        # 14) Turn OFF all layers and enable only required layers
        # ============================================================
        try:
            self.append_log("🔧 Updating layer visibility...", "info")

            project = QgsProject.instance()
            layer_tree = project.layerTreeRoot()

            # Turn OFF all layers first
            for child in layer_tree.children():
                if isinstance(child, QgsLayerTreeLayer):
                    child.setItemVisibilityChecked(False)


            # List required layer references (already loaded earlier)
            
            required_layers = [
                no_dup_layer,                 # 1
                self.parcel_new,              # 2
                self.builtup_ref_layer,       # 3
                parcel_explode_path,          # 4
                builtup_explode_path,         # 5
                parcel_clip_explode_path,     # 6
            ]
            
            # Turn ON visibility for required layers only
            for lyr in required_layers:
                if lyr is not None:
                    node = layer_tree.findLayer(lyr.id())

                    if node:
                        node.setItemVisibilityChecked(True)

                    else:
                        self.append_log(f"⚠️ Layer tree node not found for {lyr.name()}", "error")

            self.advance_progress(f"✅ Turned OFF all layers and enable only required layers", "success")

        except Exception as e:
            self.append_log(f"❌ Error Turning OFF all layers and enable only required layers", "error")
            raise

        # ============================================================
        # 15) Reorder layers (TOP → BOTTOM)
        # ============================================================
        self.append_log("🔄 Reordering layers...", "info")

        # root = QgsProject.instance().layerTreeRoot()

        try:
            root = QgsProject.instance().layerTreeRoot()

            # Reorder layers so that the first in the list becomes top
            for i, lyr in enumerate(required_layers):
                if lyr is None:
                    continue

                node = root.findLayer(lyr.id())
                if node:
                    # Clone node and insert at top position (i)
                    root.insertChildNode(i, node.clone())

                    # Remove the original node
                    root.removeChildNode(node)
                else:
                    self.append_log(f"⚠️ Layer not found in tree: {lyr.name()}", "warning")

            self.advance_progress("✅ Layers reordered successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error during layer reorder: {str(e)}", "error")

        # --- Save project again so the layout persists ---
        try:
            project_now_1 = QgsProject.instance()

            project_now_1.write(self.project_path)
            self.append_log("✅ Project saved successfully.", "success")
        except Exception as e:
            self.append_log(f"⚠️ Error re-saving QGIS project: {e}", "error")

        # ============================================================
        # Create a new QGIS Layer Theme for Village Map
        # ============================================================
        try:
            ppm_theme_name = "PP_Map_Theme"

            # Collect layer IDs used in Property Parcel map from required_layers (in same order)
            ppm_layer_ids = [lyr.id() for lyr in required_layers if lyr is not None]

            # Ask GUI to create theme
            self.create_theme(ppm_theme_name)

            self.advance_progress(f"🏷️ Map Theme '{ppm_theme_name}' created successfully.", "success")


        except Exception as e:
            self.append_log(f"❌ Error creating map theme: {str(e)}", "error")

        # ============================================================
        # 16) Find QPT Template → folder_support
        # ============================================================
        self.append_log("📄 Preparing PPM template...", "info")

        dest_qpt = os.path.join(os.path.dirname(__file__), "qpt", "INITIAL_PPM_A4_TEMPLATE_NEW.qpt")

        # =============================
        # 17) Load QPT Template
        # =============================
        # update = lambda msg, t="info": self.log_signal.emit(msg, t)
        self.append_log("📄 Loading PPM layout...", "info")

        project = QgsProject.instance()
        layout_mgr = project.layoutManager()

        # layout = QgsPrintLayout(project)


        # --- Save project again so the layout persists ---
        try:
            project_now = QgsProject.instance()
            project_now.write(self.project_path)

            self.append_log("✅ Project saved successfully.", "success")
        except Exception as e:
            self.append_log(f"⚠️ Error re-saving QGIS project: {e}", "error")

        # Emit layout to GUI thread
        # ============================================================
        # 18) Open the layout automatically
        # ============================================================
        try:

            # Send the layout and qpt path to the GUI
            self.open_layout(dest_qpt, self.parcel_new, self.builtup_ref_layer)

            self.append_log("📐 PPM layout opened successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error opening layout: {str(e)}", "error")

        # --- Save project again so the layout persists ---
        try:
            project_now = QgsProject.instance()
            project_now.write(self.project_path)
            self.append_log("✅ Project saved successfully.", "success")
        except Exception as e:
            self.append_log(f"⚠️ Error re-saving QGIS project: {e}", "error")
            
        # ----------------------------
        # Continue heavy processing...
        # ----------------------------
        self.advance_progress("✅ Initial PPM generation completed.", "success")
        
        
    def run_ppms_92(self,p):

        self.append_log(" ", "info")
        self.append_log(" ", "info")
        self.append_log("====================================================", "info")
        self.append_log(" ", "info")
        self.append_log("==  🗺️ PPM generation along with 9(2) Notices process starts  ==", "info")
        self.append_log("", "info")
        self.append_log("====================================================", "info")
        self.append_log(" ", "info")
        self.append_log(" ", "info")
        
        # Ensure Shapefiles folder exists
        shp_folder = os.path.join(self.folder_lv4, "Shapefiles")
        os.makedirs(shp_folder, exist_ok=True)

        # ---------------------------------------------------
        # Output path
        # ---------------------------------------------------   
        join_out_path = os.path.join(
            self.folder_lv4,"Shapefiles",
            f"{p['village']}_{p['lgd']}_attribute_joined.gpkg"
        )   # ✅ gpkg, not shp

        try:
            builtup_new_fields = {
                'PAR_REF_Built': QVariant.Int,
            }

            if not self.add_fields_safe(self.builtup_ref_layer, builtup_new_fields):
                raise Exception("Failed to create fields")
            
            builtup_update = []
            builtup_column_selected = p['builtup_ppn_column']
            for feat in self.builtup_ref_layer.getFeatures():
                builtup_update.append(
                    (feat.id(), {
                        "PAR_REF_Built": feat[builtup_column_selected],
                    })
                )
            if not self.update_attributes_safe(self.builtup_ref_layer, builtup_update):
                    raise Exception("Failed to update attributes")
            try:     
                layer = self.builtup_ref_layer
                provider = layer.dataProvider()

                existing_fields = [f.name() for f in layer.fields()]

                if "Ar_M2_Ref" not in existing_fields:
                    provider.addAttributes([
                        QgsField("Ar_M2_Ref", QVariant.Double, len=20, prec=3)
                    ])
                    layer.updateFields()

                ppn_field = p['builtup_ppn_column']   # or "PAR_REF"
                area_field = "Ar_M2"

                ppn_index = layer.fields().indexOf(ppn_field)
                area_index = layer.fields().indexOf(area_field)
                ref_index  = layer.fields().indexOf("Ar_M2_Ref")

                ppn_area_sum = {}

                for f in layer.getFeatures():
                    ppn = f[ppn_field]
                    area = f[area_field] or 0.0

                    if ppn not in ppn_area_sum:
                        ppn_area_sum[ppn] = 0.0

                    ppn_area_sum[ppn] += area


                if layer.isEditable():
                    layer.commitChanges()

                layer.startEditing()

                for f in layer.getFeatures():
                    ppn = f[ppn_field]
                    total_area = round(ppn_area_sum.get(ppn, 0.0), 3)

                    layer.changeAttributeValue(
                        f.id(),
                        ref_index,
                        total_area
                    )

                layer.commitChanges()
                layer.triggerRepaint()
                QgsApplication.processEvents()
                self.append_log(" 👍🏻 Area calculations completed for builtup layer", "success")
            except Exception as e:
                self.append_log("Error Adding areas of built-up")

            try:
                # -----------------------------------------
                # Transfer Ar_M2_Ref → parcel_layer_joined
                # -----------------------------------------

                src_layer = self.builtup_ref_layer
                tgt_layer = self.parcel_new

                src_ppn_field = p['builtup_ppn_column']
                tgt_ppn_field = p['ppn_column']

                src_ref_field = "Ar_M2_Ref"

                # Build lookup dictionary from builtup layer
                ppn_to_area = {}

                for f in src_layer.getFeatures():
                    # ppn = f[src_ppn_field]
                    # area_ref = f[src_ref_field]
                    ppn = self.normalize_ppn(f[src_ppn_field])
                    area_ref = f[src_ref_field] or 0.00

                    # Avoid overwriting if already captured
                    if ppn not in ppn_to_area:
                        ppn_to_area[ppn] = area_ref

                # Add Ar_M2_Ref field to parcel layer if not exists
                tgt_provider = tgt_layer.dataProvider()
                tgt_fields = [f.name() for f in tgt_layer.fields()]

                if "Ar_M2_Ref" not in tgt_fields:
                    tgt_provider.addAttributes([
                        QgsField("Ar_M2_Ref", QVariant.Double, len=20, prec=3)
                    ])
                    tgt_layer.updateFields()

                tgt_ref_index = tgt_layer.fields().indexOf("Ar_M2_Ref")

                if tgt_layer.isEditable():
                    tgt_layer.commitChanges()

                tgt_layer.startEditing()

                for f in tgt_layer.getFeatures():
                    ppn = self.normalize_ppn(f[tgt_ppn_field])

                    area_val = round(ppn_to_area.get(ppn, 0.0), 3)

                    tgt_layer.changeAttributeValue(
                        f.id(),
                        tgt_ref_index,
                        area_val
                    )

                tgt_layer.commitChanges()
                tgt_layer.triggerRepaint()
                QgsApplication.processEvents()
                self.append_log(" 👍🏻 Area calculations in built up layer added back to parcel layer", "success")

            except Exception as e:
                self.append_log(f"❌ Error transferring Ar_M2_Ref to parcel layer: {e}")

            self.append_log("🔗 Joining parcel layer with Excel table...")
            result = processing.run(
                "native:joinattributestable",
                {
                    'INPUT': self.parcel_new,
                    'FIELD': 'PAR_REF',          # field in initial_parcel_layer
                    'INPUT_2': self.excel_layer,
                    'FIELD_2': 'PPN',        # field in excel
                    'FIELDS_TO_COPY': [],    # [] = copy all
                    'METHOD': 0,             # 1 = one-to-many
                    'DISCARD_NONMATCHING': False,
                    'OUTPUT': join_out_path
                }
            )

            parcel_layer_joined = QgsVectorLayer(join_out_path, f"{p['village']}_{p['lgd']}_joined", "ogr")
            if parcel_layer_joined.isValid():
                QgsProject.instance().addMapLayer(parcel_layer_joined)
                self.append_log(f"✅ Joined layer created: {join_out_path}")
            else:
                self.append_log("❌ Joined layer invalid after processing.", "success")


            # ✅ Add new field "Sy_Nos_owners" in the joined layer
            parcel_layer_joined.startEditing()
            if "Sy_Nos_owners" not in [f.name() for f in parcel_layer_joined.fields()]:
                parcel_layer_joined.dataProvider().addAttributes([QgsField("Sy_Nos_owners", QVariant.Int)])
                parcel_layer_joined.updateFields()

            if "Card_ID" not in [f.name() for f in parcel_layer_joined.fields()]:
                parcel_layer_joined.dataProvider().addAttributes([QgsField("Card_ID", QVariant.String)])
                parcel_layer_joined.updateFields()

            idx_ppm = parcel_layer_joined.fields().indexFromName("PAR_REF")
            idx_syno = parcel_layer_joined.fields().indexFromName("Sy_Nos_owners")
            idx_cardid = parcel_layer_joined.fields().indexFromName("Card_ID")
            idx_owner = parcel_layer_joined.fields().indexFromName("Property type (Individual/Joint/Apartment/Government)")
            

            # Group by PPM and assign serial numbers
            grouped = defaultdict(list)
            for feat in parcel_layer_joined.getFeatures():
                ppm_value = feat.attribute(idx_ppm)
                grouped[ppm_value].append(feat.id())

            for ppm_value, feat_ids in grouped.items():
                for i, fid in enumerate(feat_ids, start=1):
                    parcel_layer_joined.changeAttributeValue(fid, idx_syno, i)

            # ✅ Now populate Card_ID based on CASE expression
            for feat in parcel_layer_joined.getFeatures():
                owner_val = str(feat.attribute(idx_owner) or "").upper()
                ppm_val = str(feat.attribute(idx_ppm))
                syno_val = feat.attribute(idx_syno)

                if "JOINT" in owner_val:
                    card_id = f"{ppm_val}_{syno_val}"
                else:
                    card_id = ppm_val

                parcel_layer_joined.changeAttributeValue(feat.id(), idx_cardid, card_id)

            # ✅ Add new field "Total_Owners" if not exists
            if "Total_Owners" not in [f.name() for f in parcel_layer_joined.fields()]:
                parcel_layer_joined.dataProvider().addAttributes([QgsField("Total_Owners", QVariant.Int)])
                parcel_layer_joined.updateFields()

            idx_total = parcel_layer_joined.fields().indexFromName("Total_Owners")

            # ✅ Build group dictionary again (PPM → [feat_ids])
            grouped = defaultdict(list)
            for feat in parcel_layer_joined.getFeatures():
                ppm_value = feat.attribute(idx_ppm)
                grouped[ppm_value].append(feat)

            # ✅ Assign max(Sy_Nos_owners) for each PPM to all its features
            for ppm_value, feats in grouped.items():
                max_val = max(f.attribute(idx_syno) for f in feats if f.attribute(idx_syno) is not None)
                for f in feats:
                    parcel_layer_joined.changeAttributeValue(f.id(), idx_total, max_val)

            # ✅ Find global max Total_Owners across all PPMs
            idx_total = parcel_layer_joined.fields().indexFromName("Total_Owners")
            max_total = max(
                (f.attribute(idx_total) for f in parcel_layer_joined.getFeatures() if f.attribute(idx_total) is not None),
                default=0
            )

            # Commit changes
            parcel_layer_joined.commitChanges()
            QgsApplication.processEvents()
            self.append_log(" ✅ Serial numbers and Card_ID added in joined layer.", "success")
            self.append_log(" ✅ Total_Owners column populated with max Sy_Nos_owners per PPN.", "success")
            self.append_log(f" 👍🏻 Max Total_Owners across PPMs = {max_total}", "success")

        except Exception as e:
            self.append_log(f"❌ Error in creating Joined layer : {e}", "error")
            raise 

        # ------------------------------------------
        # 1. Boundary from parcel shapefile & builtup shapefile
        # Output: Parcel_Boundary.shp & Builtup_Boundary.shp in folder_shp
        # ------------------------------------------
        try:
            parcel_boundary_path = os.path.join(shp_folder, "Parcel_Boundary.shp")
            params_parcel_boundary = {
                'INPUT': parcel_layer_joined,
                'OUTPUT': parcel_boundary_path
            }
            self.append_log("⌛ Running Parcel Boundary algorithm...", "info")
            res_boundary = processing.run("native:boundary", params_parcel_boundary)
            self.advance_progress(f"📤 Parcel Boundary created: {parcel_boundary_path}", "success")

            builtup_boundary_path = os.path.join(shp_folder, "Builtup_Boundary.shp")
            params_builtup_boundary = {
                'INPUT': self.builtup_ref_layer,
                'OUTPUT': builtup_boundary_path
            }
            self.append_log("⌛ Running Built-up Boundary algorithm...", "info")
            res_boundary = processing.run("native:boundary", params_builtup_boundary)
            self.advance_progress(f"📤 Built-up Boundary created: {builtup_boundary_path}", "success")
        except Exception as e:
            self.append_log(f"❌ Error creating Parcel Boundary & Built-up Boundary: {e}", "error")

        # ------------------------------------------
        # 2. Explode lines 
        # ------------------------------------------
        try:
            # ------------------------------------------
            # 2A. Explode lines (Parcel_Boundary -> Parcel_Explode_Lines.shp)
            # ------------------------------------------
            parcel_explode = os.path.join(shp_folder, "Parcel_Explode_Lines.shp")
            params_explode1 = {
                'INPUT': parcel_boundary_path,
                'OUTPUT': parcel_explode
            }
            self.append_log("⌛ Exploding Parcel boundary to lines...", "info")
            res_explode1 = processing.run("native:explodelines", params_explode1)
            self.advance_progress(f"📤 Exploded Parcel lines exported: {parcel_explode}", "success")
            parcel_explode_layer_name = "Parcel_Explode_Lines"
            parcel_explode_path = self.readding_layer(parcel_explode_layer_name, parcel_explode)

            # ------------------------------------------
            # 2B. Explode lines (Builtup_Boundary -> Builtup_Explode_Lines.shp)
            # ------------------------------------------
            builtup_explode = os.path.join(shp_folder, "Builtup_Explode_Lines.shp")
            params_explode = {
                'INPUT': builtup_boundary_path,
                'OUTPUT': builtup_explode
            }
            self.append_log("⌛ Exploding Builtup boundary to lines...", "info")
            res_explode1 = processing.run("native:explodelines", params_explode)
            self.advance_progress(f"📤 Exploded Builtup lines exported: {builtup_explode}", "success")
            builtup_explode_layer_name = "Builtup_Explode_Lines"
            builtup_explode_path = self.readding_layer(builtup_explode_layer_name, builtup_explode)

        except Exception as e:
            self.append_log(f"❌ Error creating Explode lines for Parcel Boundary & Built-up Boundary: {e}", "error")


        try:
            # ------------------------------------------
            # 3. Extract vertices from Parcel_Boundary -> Parcel_Vertices.shp
            # ------------------------------------------
            parcel_vertices_path = os.path.join(shp_folder, "Parcel_Vertices.shp")
            params_vertices = {
                'INPUT': parcel_boundary_path,
                'OUTPUT': parcel_vertices_path
            }
            self.append_log("⌛ Extracting vertices...", "info")
            res_vertices = processing.run("native:extractvertices", params_vertices)
            self.advance_progress(f"📤 Vertices extracted: {parcel_vertices_path}", "success")
        except Exception as e:
            self.append_log(f"❌ Error creating vertices for Parcel layer: {e}", "error")


        # ------------------------------------------
        # 4. Multi-ring buffer around vertices
        #    (NUMBER=1, DISTANCE=2, SEGMENTS=8) -> Parcel_Buffer.shp
        # Note: native:multiringbuffer expects a comma-separated DISTANCES string
        # ------------------------------------------

        parcel_buffer_path = os.path.join(shp_folder, "Parcel_Buffer.shp")

        self.append_log("⌛ Creating multi-ring buffer around vertices...", "info")
        try:
            # Try native:multiringbuffer first
            params_buffer = {
                'INPUT': parcel_vertices_path,
                'DISTANCE': 2,   # distance per ring
                'NUMBER': 1,     # number of rings
                'SEGMENTS': 8,
                'OUTPUT': parcel_buffer_path
            }
            res_buffer = processing.run("native:multiringbuffer", params_buffer)
            self.advance_progress(f"📤 Multi-ring buffer created: {parcel_buffer_path}", "success")
        except Exception as e1:
            # self.append_log(f"⚠️ native:multiringbuffer failed ({str(e1)}), trying native:buffer...", "error")
            try:
                # 2️⃣ Fallback to simple buffer
                params_fallback = {
                    'INPUT': parcel_vertices_path,
                    'DISTANCE': 1,
                    'SEGMENTS': 8,
                    'DISSOLVE': False,
                    'OUTPUT': parcel_buffer_path
                }
                res = processing.run("native:buffer", params_fallback)
                self.advance_progress(f"📤 Buffer created using fallback: {parcel_buffer_path}", "success")

            except Exception as e2:
                # self.append_log(f"⚠️ native:buffer also failed ({str(e2)}), trying native:multiringconstantbuffer...", "error")
                # 3️⃣ Last fallback: native:multiringconstantbuffer
                params_constant = {
                    'INPUT': parcel_vertices_path,
                    'RINGS': 1,
                    'DISTANCE': 1,
                    'OUTPUT': parcel_buffer_path
                }
                res = processing.run("native:multiringconstantbuffer", params_constant)
                self.advance_progress(f"📤 Multi-ring constant buffer created: {parcel_buffer_path}", "success")

        # ------------------------------------------
        # 5. Clip: parcel_explode (input) clipped by parcel_buffer (overlay) -> Parcel_Clip.shp
        # ------------------------------------------
        try:
            parcel_clip_path = os.path.join(shp_folder, "Parcel_Clip.shp")
            params_clip = {
                'INPUT': parcel_explode_path,
                'OVERLAY': parcel_buffer_path,
                'OUTPUT': parcel_clip_path
            }
            self.append_log("⌛ Clipping exploded lines by buffer...", "info")
            res_clip = processing.run("native:clip", params_clip)
            self.advance_progress(f"📤 Clipped layer created: {parcel_clip_path}", "success")
        except Exception as e:
            self.append_log(f"❌ Error creating Clip layer: {e}", "error")

        # ------------------------------------------
        # 6. Explode lines for the clip shapefile -> Parcel_Clip_Explode_Lines.shp
        # ------------------------------------------
        try:
            parcel_clip_explode = os.path.join(shp_folder, "Parcel_Clip_Explode_Lines.gpkg")
            params_explode2 = {
                'INPUT': parcel_clip_path,
                'OUTPUT': parcel_clip_explode
            }
            self.append_log("⌛ Exploding clipped lines...", "info")
            res_explode2 = processing.run("native:explodelines", params_explode2)
            self.advance_progress(f"📤 Exploded clipped lines exported: {parcel_clip_explode}", "success")
            Parcel_Clip_Explode_Lines_layer_name = "Parcel_Clip_Explode_Lines"
            parcel_clip_explode_path = self.readding_layer(Parcel_Clip_Explode_Lines_layer_name, parcel_clip_explode)
        except Exception as e:
            self.append_log(f"❌ Error creating Explode lines for the clip shapefile: {e}", "error")

        # ------------------------------------------
        # 7A. Add 'length' field to Parcel_Explode_Lines and populate with geometry length
        # ------------------------------------------
        self.append_log("⌛ Adding 'length' field for Parcel and populating values...", "info")
        
        try:
            # load exploded lines layer (first explode result)
            # Parcel_explode_lines_layer = QgsVectorLayer(parcel_explode_path, "Parcel_Explode_Lines", "ogr")
            Parcel_explode_lines_layer = parcel_explode_path

            if not Parcel_explode_lines_layer.isValid():
                self.append_log("❌ Failed to open Parcel exploded lines layer for length calculation.", "error")
            else:
                dp = Parcel_explode_lines_layer.dataProvider()
                # Add field if possible
                if dp.capabilities() & dp.AddAttributes:
                    Parcel_explode_lines_layer.startEditing()
                    fld = QgsField("length", QVariant.Double, "", 10, 2)
                    dp.addAttributes([fld])
                    Parcel_explode_lines_layer.updateFields()

                    idx = Parcel_explode_lines_layer.fields().indexFromName("length")
                    # iterate and set length
                    for feat in Parcel_explode_lines_layer.getFeatures():
                        geom = feat.geometry()
                        if geom is None:
                            continue
                        length_value = geom.length()
                        Parcel_explode_lines_layer.changeAttributeValue(feat.id(), idx, round(float(length_value), 2))
                    Parcel_explode_lines_layer.commitChanges()
                    self.advance_progress("✅ 'length' field added and populated on Parcel_Explode_Lines.", "success")
                else:
                    self.append_log("❌ Layer provider does not support adding attributes.", "error")
        except Exception as e:
            self.append_log(f"❌ Error Adding 'length' field to parcels : {e}", "error")

        # ------------------------------------------
        # 7B. Add 'length' field to Builtup_Explode_Lines and populate with geometry length
        # ------------------------------------------
        try:
            self.append_log("⌛ Adding 'length' field for Builtup and populating values...", "info")

            # load exploded lines layer (first explode result)
            # Builtup_explode_lines_layer = QgsVectorLayer(builtup_explode_path, "Builtup_Explode_Lines", "ogr")
            Builtup_explode_lines_layer = builtup_explode_path

            if not Builtup_explode_lines_layer.isValid():
                self.append_log("❌ Failed to open Builtup exploded lines layer for length calculation.", "error")
            else:
                dp = Builtup_explode_lines_layer.dataProvider()
                # Add field if possible
                if dp.capabilities() & dp.AddAttributes:
                    Builtup_explode_lines_layer.startEditing()
                    fld = QgsField("length", QVariant.Double, "", 10, 2)
                    dp.addAttributes([fld])
                    Builtup_explode_lines_layer.updateFields()

                    idx = Builtup_explode_lines_layer.fields().indexFromName("length")
                    # iterate and set length
                    for feat in Builtup_explode_lines_layer.getFeatures():
                        geom = feat.geometry()
                        if geom is None:
                            continue
                        length_value = geom.length()
                        Builtup_explode_lines_layer.changeAttributeValue(feat.id(), idx, round(float(length_value), 2))
                    Builtup_explode_lines_layer.commitChanges()
                    self.advance_progress("✅ 'length' field added and populated on Builtup_Explode_Lines.", "success")
                else:
                    self.append_log("❌ Layer provider does not support adding attributes.", "error")
        except Exception as e:
            self.append_log(f"❌ Error Adding 'length' field to built-up : {e}", "error")

        # ============================================================
        # 8. Add fields (point_ID, Easting_X, Northing_Y) to vertices
        # ============================================================
        try:
            self.append_log("⌛ Adding point_ID, Easting_X, Northing_Y fields to vertices...", "info")

            vertices_layer = QgsVectorLayer(parcel_vertices_path, "Parcel_Vertices", "ogr")
            if not vertices_layer.isValid():
                self.append_log("❌ Failed to load Parcel_Vertices for attribute update.", "error")
            else:
                dp = vertices_layer.dataProvider()

                if dp.capabilities() & QgsVectorDataProvider.AddAttributes:
                    vertices_layer.startEditing()
                    fdefs = []

                    if vertices_layer.fields().indexFromName("point_ID") == -1:
                        fdefs.append(QgsField("point_ID", QVariant.Int, '', 10, 0))
                    if vertices_layer.fields().indexFromName("Easting_X") == -1:
                        fdefs.append(QgsField("Easting_X", QVariant.Double, '', 15, 3))
                    if vertices_layer.fields().indexFromName("Northing_Y") == -1:
                        fdefs.append(QgsField("Northing_Y", QVariant.Double, '', 15, 3))

                    if fdefs:
                        dp.addAttributes(fdefs)

                    vertices_layer.updateFields()

                    idx_pid = vertices_layer.fields().indexFromName("point_ID")
                    idx_e = vertices_layer.fields().indexFromName("Easting_X")
                    idx_n = vertices_layer.fields().indexFromName("Northing_Y")

                    has_vertex_ind = vertices_layer.fields().indexFromName("vertex_ind") != -1

                    for i, feat in enumerate(vertices_layer.getFeatures()):
                        geom = feat.geometry()
                        if geom is None:
                            continue

                        pt = geom.asPoint()  # Extract Vertices ALWAYS returns point geometry

                        # point_ID logic
                        if idx_pid != -1:
                            if has_vertex_ind:
                                try:
                                    vertices_layer.changeAttributeValue(feat.id(), idx_pid,
                                                                        int(feat.attribute("vertex_ind")) + 1)
                                except Exception:
                                    vertices_layer.changeAttributeValue(feat.id(), idx_pid, i + 1)
                            else:
                                vertices_layer.changeAttributeValue(feat.id(), idx_pid, i + 1)

                        # Easting/ Northing
                        if pt:
                            if idx_e != -1:
                                vertices_layer.changeAttributeValue(feat.id(), idx_e, round(pt.x(), 3))
                            if idx_n != -1:
                                vertices_layer.changeAttributeValue(feat.id(), idx_n, round(pt.y(), 3))

                    vertices_layer.commitChanges()
                    self.advance_progress("✅ Fields added & populated for Parcel_Vertices.", "success")
                else:
                    self.append_log("❌ Provider does not support attribute additions on vertices.", "error")
        except Exception as e:
            self.append_log(f"❌ Error Adding 'co-ordinates' : {e}", "error")

        # ============================================================
        # 9–11. Remove duplicate vertices (PAR_REF, Easting_X, Northing_Y)
        # Robust version: use geometry coords (not relying on attr columns)
        # ============================================================

        try:
            self.append_log("⌛ Removing duplicate vertices...", "info")

            if not vertices_layer or not vertices_layer.isValid():
                raise Exception("Vertices layer is not loaded or invalid.")

            # --------------------------------------------------------
            # 1 ) Run duplicate removal
            # --------------------------------------------------------
            no_dup_path_link = os.path.join(shp_folder, "Parcel_No_Dup_Vertices.shp")
            if os.path.exists(no_dup_path_link):
                os.remove(no_dup_path_link)

            result = processing.run(
                "native:removeduplicatesbyattribute",
                {
                    'INPUT': vertices_layer,
                    'FIELDS': ['PAR_REF', 'Easting_X', 'Northing_Y'],
                    'OUTPUT': no_dup_path_link
                }
            )
            # --------------------------------------------------------
            # 2 ) Reload final layer to project
            # --------------------------------------------------------
            no_dup_layer = self.readding_layer("Parcel_No_Dup_Vertices", no_dup_path_link)
            if not no_dup_layer:
                raise Exception("Failed to load the no-duplicate output layer.")

            self.advance_progress("✅ Duplicate vertices removed & saved successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error removing duplicate vertices: {e}", "error")
            raise

        # ============================================================
        # 13 ) Apply QML Style Files
        # ============================================================
        try:
            self.append_log("🎨 Applying QML styles...", "info")

            style_base_9_2 = os.path.join(
                os.path.dirname(__file__),
                "styling_properties",
                "9(2)"
            )

            # QML mappings
            vm_qml_files = {
                parcel_layer_joined: "Parcel_Polygon.qml",
                parcel_explode_path: "Parcel_explode_lines.qml",
                builtup_explode_path: "Builtup_explode_lines.qml",
                parcel_clip_explode_path: "Parcel_clip_explode_lines.qml",
                no_dup_layer: "Parcel_no_dup_vertices.qml",
                self.builtup_ref_layer: "Builtup_Ref.qml"
            }

            # Apply styling
            for layer_name, qml_file in vm_qml_files.items():
                qml_file_path  = os.path.join(style_base_9_2, qml_file)
                self.apply_qml_style(layer_name, qml_file_path)

            self.advance_progress(f"✅ QML Style Applied to all required fiels", "success")

            
        except Exception as e:
            self.append_log(f"❌ Error Applying QML styles", "error")
            raise

        # ============================================================
        # Turn OFF all layers and enable only required layers
        # ============================================================
        try:
            self.append_log("🔧 Updating layer visibility...", "info")
            
            project = QgsProject.instance()
            layer_tree = project.layerTreeRoot()

            # Turn OFF all layers first
            for child in layer_tree.children():
                if isinstance(child, QgsLayerTreeLayer):
                    
                    child.setItemVisibilityChecked(False)

            required_layer_now = [no_dup_layer, parcel_layer_joined, self.builtup_ref_layer, parcel_explode_path, builtup_explode_path, parcel_clip_explode_path]
            

            # Turn ON visibility for required layers only
            for lyr in required_layer_now:
                if lyr is not None:
                    node = layer_tree.findLayer(lyr.id())
                    if node:
                        node.setItemVisibilityChecked(True)

                    else:
                        self.append_log(f"⚠️ Layer tree node not found for {lyr.name()}", "warning")
            self.advance_progress(f"✅ Turned OFF all layers and enable only required layers", "success")
        except Exception as e:
            self.append_log(f"❌ Error Turning OFF all layers and enable only required layers", "error")
            raise

        # ============================================================
        # Reorder layers (TOP → BOTTOM)
        # ============================================================
        self.append_log("🔄 Reordering layers...", "info")

        try:
            root = QgsProject.instance().layerTreeRoot()

            # Reorder layers so that the first in the list becomes top
            for i, lyr in enumerate(required_layer_now):
                if lyr is None:
                    continue

                node = root.findLayer(lyr.id())

                if node:
                    # Clone node and insert at top position (i)
                    root.insertChildNode(i, node.clone())

                    # Remove the original node
                    root.removeChildNode(node)
                else:
                    self.append_log(f"⚠️ Layer not found in tree: {lyr.name()}", "warning")

            self.advance_progress("✅ Layers reordered successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error during layer reorder: {str(e)}", "error")
            raise

        # --- Save project again so the layout persists ---
        try:
            project_now_4 = QgsProject.instance()

            project_now_4.write(self.project_path)
            self.append_log("✅ Project saved successfully.", "success")
        except Exception as e:
            self.append_log(f"⚠️ Error re-saving QGIS project: {e}", "error")

        # ============================================================
        # Create a new QGIS Layer Theme for Village Map
        # ============================================================
        try:
            theme_name = "PPMs with 9(2) Notices"

            # # Collect layer IDs used in village map
            # vm_layer_ids = [parcel_vm.id(), parcel_vm_dissolve.id()]

            # Ask GUI to create theme
            self.create_theme(theme_name)

            self.advance_progress(f"🏷️ Map Theme '{theme_name}' created successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error creating map theme: {str(e)}", "error")

        # ============================================================
        # Find QPT Template → folder_support
        # ============================================================
        self.append_log("📄 Preparing PPNs Map with 9(2) Notices template...", "info")

        dest_qpt_9_2 = os.path.join(os.path.dirname(__file__), "qpt", "PPM_With_9(2)_Notice.qpt")

        # =============================
        # Load QPT Template
        # =============================
        # update = lambda msg, t="info": self.log_signal.emit(msg, t)
        self.append_log("📄 Loading PPNs Map with 9(2) Notices template layout...", "info")

        project = QgsProject.instance()
        layout_mgr = project.layoutManager()
        
        # --- Save project again so the layout persists ---
        try:
            project_now_9_2 = QgsProject.instance()
            project_now_9_2.write(self.project_path)

            self.append_log("✅ Project saved successfully with templete.", "success")
        except Exception as e:
            self.append_log(f"⚠️ Error re-saving QGIS project: {e}", "error")

        # Emit layout to GUI thread
        # ============================================================
        # 18) Open the layout automatically
        # ============================================================
        try:

            # Send the layout and qpt path to the GUI
            self.open_9_2_layout(dest_qpt_9_2, parcel_layer_joined)

            self.advance_progress("📐 PPNs Map with 9(2) Notices layout opened successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error opening layout: {str(e)}", "error")

        # --- Save project again so the layout persists ---
        try:
            project_now_5 = QgsProject.instance()
            project_now_5.write(self.project_path)
            self.append_log("✅ Project saved successfully.", "success")
        except Exception as e:
            self.append_log(f"⚠️ Error re-saving QGIS project: {e}", "error")

        # ------------------------------------------
        # FINISH
        # ------------------------------------------

        self.advance_progress("✅ PPNs Map with 9(2) Notices Layout generation completed successfully.", "success")


    def run_village_map(self,p):

        self.append_log(" ", "info")
        self.append_log(" ", "info")
        self.append_log("====================================================", "info")
        self.append_log(" ", "info")
        self.append_log("========  🗺️ Village Map generation process starts  ========", "info")
        self.append_log("", "info")
        self.append_log("====================================================", "info")
        self.append_log(" ", "info")
        self.append_log(" ", "info")
        
        parcel_new_out = os.path.join(self.folder_shp, "Parcel_VM.gpkg")
        parcel_vm = self.safe_export_layer(self.parcel_new, parcel_new_out, "Parcel_VM")

        try:
            self.calculate_total_area_fields(parcel_vm)
            self.advance_progress("📊 Total area fields (ToT_Ar_M2, ToT_Ar_Y2) calculated.", "success")
        except Exception as e:
            self.append_log("❌ Error in calculating total areas")
        
        # -----------------------------------------------------------
        # Dissolving Parcel_Boundary for village outer boundary
        # -----------------------------------------------------------
        try:
            parcel_dissolve_vm = os.path.join(self.folder_shp, "Parcel_dissolve_vm.gpkg")
            params_dissolve = {
                'INPUT': parcel_vm,
                # 'FIELD':['DISTRICT','MANDAL','VILLAGE','LGD_CODE','Ar_M2','Ar_Ya2'],
                'OUTPUT': parcel_dissolve_vm
            }

            self.append_log("⌛ Extracting vertices...", "info")
            
            res_dissolve = processing.run("native:dissolve", params_dissolve)
            self.append_log(f"📤 Parcel Dissolved: {parcel_dissolve_vm}", "success")

            vm_dissolve_name = "Parcel_No_Dup_Vertices"
            parcel_vm_dissolve = self.readding_layer(vm_dissolve_name, parcel_dissolve_vm)

            self.advance_progress("🔄 Parcel Dissolved layer added back into the project.", "info")
        except Exception as e:
            self.append_log(f"❌ Error Dissolving Parcel_Boundary for village outer boundary", "error")
            raise
        
        # ============================================================
        # Turn OFF all layers and enable only required layers
        # ============================================================
        try:
            self.append_log("🔧 Updating layer visibility...", "info")
            
            project = QgsProject.instance()
            layer_tree = project.layerTreeRoot()

            # Turn OFF all layers first
            for child in layer_tree.children():
                if isinstance(child, QgsLayerTreeLayer):
                    
                    child.setItemVisibilityChecked(False)

            required_layer_now = [parcel_vm, parcel_vm_dissolve]
            

            # Turn ON visibility for required layers only
            for lyr in required_layer_now:
                if lyr is not None:
                    node = layer_tree.findLayer(lyr.id())
                    if node:
                        node.setItemVisibilityChecked(True)

                    else:
                        self.append_log(f"⚠️ Layer tree node not found for {lyr.name()}", "warning")
            self.advance_progress(f"✅ Turned OFF all layers and enable only required layers", "success")
        except Exception as e:
            self.append_log(f"❌ Error Turning OFF all layers and enable only required layers", "error")
            raise

        # ============================================================
        # Reorder layers (TOP → BOTTOM)
        # ============================================================
        self.append_log("🔄 Reordering layers...", "info")

        try:
            root = QgsProject.instance().layerTreeRoot()

            # Reorder layers so that the first in the list becomes top
            for i, lyr in enumerate(required_layer_now):
                if lyr is None:
                    continue

                node = root.findLayer(lyr.id())

                if node:
                    # Clone node and insert at top position (i)
                    root.insertChildNode(i, node.clone())

                    # Remove the original node
                    root.removeChildNode(node)
                else:
                    self.append_log(f"⚠️ Layer not found in tree: {lyr.name()}", "warning")

            self.advance_progress("✅ Layers reordered successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error during layer reorder: {str(e)}", "error")
            raise

        # --- Save project again so the layout persists ---
        try:
            project_now_2 = QgsProject.instance()

            project_now_2.write(self.project_path)
            self.append_log("✅ Project saved successfully.", "success")
        except Exception as e:
            self.append_log(f"⚠️ Error re-saving QGIS project: {e}", "error")

        # ============================================================
        # Apply QML Style Files
        # ============================================================
        try:
            self.append_log("🎨 Applying QML styles...", "info")
            style_base_vm = os.path.join(
                os.path.dirname(__file__),
                "styling_properties",
                "Village Map"
            )

            # QML mappings
            vm_qml_files = {
                parcel_vm: "Parcel_village_map.qml",
                parcel_vm_dissolve: "Parcel_vm_dissolve.qml"
            }

            # Apply styling
            for layer_name, qml_file in vm_qml_files.items():
                qml_file_path  = os.path.join(style_base_vm, qml_file)
                self.apply_qml_style(layer_name, qml_file_path)

            self.advance_progress(f"✅ QML Style Applied to all required fiels", "success")

        except Exception as e:
            self.append_log(f"❌ Error Applying QML styles", "error")
            raise
        
        # ============================================================
        # Create a new QGIS Layer Theme for Village Map
        # ============================================================
        try:
            theme_name = "Village_Map_Theme"

            # Collect layer IDs used in village map
            vm_layer_ids = [parcel_vm.id(), parcel_vm_dissolve.id()]

            # Ask GUI to create theme
            self.create_theme(theme_name)

            self.advance_progress(f"🏷️ Map Theme '{theme_name}' created successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error creating map theme: {str(e)}", "error")
            
        # ============================================================
        # Find QPT Template → folder_support
        # ============================================================
        self.append_log("📄 Preparing SVAMITVA Village A0 Map template...", "info")

        vm_dest_qpt = os.path.join(os.path.dirname(__file__), "qpt", "PPM_VM_A0_TEMPLATE_NEW.qpt")

        # =============================
        # Load QPT Template
        # =============================
        # update = lambda msg, t="info": self.log_signal.emit(msg, t)
        self.append_log("📄 Loading SVAMITVA Village A0 map layout...", "info")

        project = QgsProject.instance()
        layout_mgr = project.layoutManager()
        
        # --- Save project again so the layout persists ---
        try:
            project_now_vm = QgsProject.instance()
            project_now_vm.write(self.project_path)

            self.append_log("✅ Project saved successfully with templete.", "success")
        except Exception as e:
            self.append_log(f"⚠️ Error re-saving QGIS project: {e}", "error")

        # Emit layout to GUI thread
        # ============================================================
        # 18) Open the layout automatically
        # ============================================================
        try:

            # Send the layout and qpt path to the GUI
            self.open_vm_layout(vm_dest_qpt, parcel_vm, parcel_vm_dissolve)

            self.advance_progress("📐 Village Map layout opened successfully.", "success")

        except Exception as e:
            self.append_log(f"❌ Error opening layout: {str(e)}", "error")

        # --- Save project again so the layout persists ---
        try:
            project_now_3 = QgsProject.instance()
            project_now_3.write(self.project_path)
            self.append_log("✅ Project saved successfully.", "success")
        except Exception as e:
            self.append_log(f"⚠️ Error re-saving QGIS project: {e}", "error")

        # ------------------------------------------
        # FINISH
        # ------------------------------------------

        self.append_log("✅ Village Map generation completed successfully.", "success")
