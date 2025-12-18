from qgis.PyQt.QtWidgets import QDialog, QFileDialog, QMessageBox
from qgis.PyQt.QtCore import QVariant
from qgis.core import (
    QgsProject, QgsVectorLayer, QgsField, QgsFeature, QgsVectorFileWriter,
    QgsWkbTypes, QgsFeatureRequest, QgsSpatialIndex, QgsLayerTreeLayer
)
import math
import processing
from .reprocess_dialog import Ui_Dialog
from PyQt5 import QtWidgets
import os

class ReprocessDialog(QDialog, Ui_Dialog):
    def __init__(self, parent=None):
        super(ReprocessDialog, self).__init__(parent)
        self.setupUi(self)
        self.init_ui()
        
    def init_ui(self):
        """Initialize UI elements and signals"""
        self.fill_polygon_layers()

        # Signals
        self.cmb_polygon_layers.currentIndexChanged.connect(self.layer_changed)
        self.cmb_abr_column.currentIndexChanged.connect(self.abr_column_changed)
        
        self.btn_check_abr.clicked.connect(self.check_abr_blanks)
        self.btn_browse_output.clicked.connect(self.select_output_folder)
        self.btn_run.clicked.connect(self.run_processing)
        self.btn_cancel.clicked.connect(self.close)

        # Fill District dropdown
        self.comboDistrict.addItems([
            "-- Select District --",
            "Alluri Sitarama Raju", "Anakapalli", "Anantapur", "Annamayya", "Bapatla", "Chittoor", "Dr. B.R. Ambedkar Konaseema", "East Godavari", "Eluru", "Guntur",
            "Kakinada", "Krishna", "Kurnool", "Nandyal", "NTR", "Palnadu", "Parvathipuram Manyam", "Prakasam", "Sri Potti Sriramulu Nellore",
            "Sri Satya Sai", "Srikakulam", "Tirupati", "Visakhapatnam", "Vizianagaram",
            "West Godavari", "YSR Kadapa"
        ])        

    def abr_column_changed(self):
        """Triggered whenever ABR column dropdown changes"""
        self.update_open_built_values()
        self.check_abr_blanks(auto=False)


    def log(self, message):
        self.txt_log.append(message)

    # --------------------------------------------------------------------
    # LOAD POLYGON LAYERS
    # --------------------------------------------------------------------
    def fill_polygon_layers(self):
        self.cmb_polygon_layers.clear()

        layers = [
            lyr for lyr in QgsProject.instance().mapLayers().values()
            if lyr.geometryType() == 2  # polygon
        ]

        self.polygon_layers = layers

        for lyr in layers:
            self.cmb_polygon_layers.addItem(lyr.name())

        if layers:
            self.layer_changed()

    # --------------------------------------------------------------------
    # UPDATE ALL DROPDOWNS WHEN LAYER CHANGES
    # --------------------------------------------------------------------
    def layer_changed(self):
        self.update_abr_columns()
        self.update_open_built_values()

    # --------------------------------------------------------------------
    # UPDATE COLUMN DROPDOWN
    # --------------------------------------------------------------------
    def update_abr_columns(self):
        self.cmb_abr_column.clear()

        idx = self.cmb_polygon_layers.currentIndex()
        if idx < 0:
            return

        layer = self.polygon_layers[idx]
        fields = [f.name() for f in layer.fields()]
        
        # Add default placeholder item
        self.cmb_abr_column.addItem("-- Select ABR Column --")

        self.cmb_abr_column.addItems(fields)

    # --------------------------------------------------------------------
    # POPULATE OPEN AND BUILT VALUES
    # --------------------------------------------------------------------
    def update_open_built_values(self):
        self.cmb_open_val.clear()
        self.cmb_built_val.clear()

        idx = self.cmb_polygon_layers.currentIndex()
        if idx < 0:
            return
        
        layer = self.polygon_layers[idx]


        abr_col = self.cmb_abr_column.currentText()

        # If default placeholder selected → stop
        if abr_col == "-- Select ABR Column --":
            return

        values = set()
        for f in layer.getFeatures():
            val = f[abr_col]
            if val not in [None, "", " "]:
                values.add(str(val))

        if len(values) == 0:
            # fully blank column → dropdowns remain empty
            self.log("ABR column is fully blank → no values for dropdowns")
            return
        
        # Fill both dropdowns with same values
        val_list = sorted(values)
        self.cmb_open_val.addItems(val_list)
        self.cmb_built_val.addItems(val_list)

    # --------------------------------------------------------------------
    # CHECK BLANK ABR VALUES
    # --------------------------------------------------------------------
    def check_abr_blanks(self, auto=False):
        idx_layer = self.cmb_polygon_layers.currentIndex()
        idx_field = self.cmb_abr_column.currentIndex()

        if idx_layer < 0 or idx_field <= 0:  # 0 = placeholder
            if not auto:
                QMessageBox.warning(self, "Warning", "Select a valid ABR column")
            return

        layer = self.polygon_layers[idx_layer]
        field_name = self.cmb_abr_column.currentText()

        blanks = 0
        unique_vals = set()

        for f in layer.getFeatures():
            v = f[field_name]

            # Treat ALL of these as blank
            if (
                v is None or
                str(v).strip() == "" or
                str(v).lower().strip() == "null" or
                str(v).lower().strip() == "<null>" or
                str(v).lower().strip() == "none"
            ):
                blanks += 1
            else:
                unique_vals.add(str(v))

        # Prepare result text
        info = f"Blanks: {blanks}\nUnique Values: {len(unique_vals)}\n"
        if len(unique_vals) > 0:
            info += "Values: " + ", ".join(sorted(unique_vals))

        # → Always show popup for both auto and manual modes
        # Because you requested auto-trigger popup

        # Case 1: Column fully blank
        if len(unique_vals) == 0:
            QMessageBox.warning(self, "ABR Status",
                                "Selected column is fully blank.\n"
                                "No valid values found.")
            return

        # Case 2: More than 2 values
        if len(unique_vals) > 2:
            QMessageBox.critical(self, "Invalid ABR Data",
                                f"There should be only 2 values (Built & Open)\n\n"
                                f"But found {len(unique_vals)} values:\n"
                                f"{', '.join(sorted(unique_vals))}")
            return

        # Case 3: Good data
        QMessageBox.information(self, "ABR Check", info)



    # --------------------------------------------------------------------
    # BROWSE FOLDER
    # --------------------------------------------------------------------
    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.txt_output.setText(folder)

    # --------------------------------------------------------------------
    # MAIN PROCESSING STARTS HERE
    # --------------------------------------------------------------------
    def run_processing(self):

        # --------------------------
        # Switch to Log tab
        # --------------------------        
        self.txt_log.clear() # Clear log
        self.dlg_reprocess.setCurrentIndex(1)  # 1 = Log tab
        
        # Input validation
        """Run the processing after validating mandatory fields"""

        # --------------------------
        # Mandatory field validation
        # --------------------------
        errors = []

        if self.cmb_polygon_layers.currentIndex() < 0:
            errors.append("Polygon Layer")
        if self.cmb_abr_column.currentIndex() <= 0:  # 0 = placeholder
            errors.append("ABR Column")
        if not self.txt_output.text().strip():
            errors.append("Output Folder")
        if self.comboDistrict.currentText() == "-- Select District --":
            errors.append("District")
        if not self.lineMandal.text().strip():
            errors.append("Mandal")
        if not self.lineVillage.text().strip():
            errors.append("Revenue Village")
        if not self.lineLGD.text().strip():
            errors.append("Revenue LGD Code")

        if errors:
            QMessageBox.warning(
                self,
                "Missing Mandatory Fields",
                "Please fill/select the following fields:\n- " + "\n- ".join(errors)
            )
            return
        # --------------------------
        # Read input values
        # --------------------------
        layer = self.polygon_layers[self.cmb_polygon_layers.currentIndex()]
        abr_col = self.cmb_abr_column.currentText()
        built_val = self.cmb_built_val.currentText()
        open_val = self.cmb_open_val.currentText()
        out_folder = self.txt_output.text()
        district = self.comboDistrict.currentText()
        mandal = self.lineMandal.text().strip()
        village = self.lineVillage.text().strip()
        lgd_code = self.lineLGD.text().strip()

        # --------------------------
        # Create folder structure
        # --------------------------
        village_folder_name = f"{lgd_code}_{village}"
        base_path = os.path.join(out_folder, district, mandal, village_folder_name)

        backup_path = os.path.join(base_path, "Backup files")
        reprocessed_path = os.path.join(base_path, "Shapefiles (Reprocessed)")
        other_shapes_path = os.path.join(backup_path, "Other Shape files")  # new folder

        for folder in [base_path, backup_path, reprocessed_path, other_shapes_path]:
            if not os.path.exists(folder):
                os.makedirs(folder)
                self.log(f"Created folder: {folder}")
            else:
                self.log(f"Folder already exists: {folder}")

        # --------------------------
        # Logging
        # --------------------------      
        self.log("Folder structure ready.")  
        self.log("Processing started…")
        self.log(f"Layer: {layer.name()}")
        self.log(f"ABR Column: {abr_col}")
        self.log(f"Built Value: {built_val}")
        self.log(f"Open Value: {open_val}")
        self.log(f"District: {district}")
        self.log(f"Mandal: {mandal}")
        self.log(f"Village: {village}")
        self.log(f"LGD Code: {lgd_code}")
        self.log(f"Output Folder: {out_folder}")

        # TODO: add your shapefile regeneration here
        # ----------------------------------------
        # -------------------- Save Project Backup --------------------
        proj_path = os.path.join(backup_path, f"{village_folder_name} (Reprocessed).qgz")
        QgsProject.instance().write(proj_path)
        self.log(f"Project saved to backup: {proj_path}")

        # -------------------- Export Original Layer --------------------
        idx_layer = self.cmb_polygon_layers.currentIndex()
        src_layer = self.polygon_layers[idx_layer]

        # Original exported shapefile name must be the REAL layer name
        orig_shp_path = os.path.join(other_shapes_path, f"{src_layer.name()}.shp")

        save_opts = QgsVectorFileWriter.SaveVectorOptions()
        save_opts.driverName = "ESRI Shapefile"

        error = QgsVectorFileWriter.writeAsVectorFormatV2(
            src_layer, orig_shp_path, QgsProject.instance().transformContext(),
            save_opts
        )
        if error[0] != QgsVectorFileWriter.NoError:
            self.log(f"Error exporting original shapefile: {error}")
            return
        self.log(f"Original shapefile exported to: {orig_shp_path}")

        # Load exported shapefile as working layer
        working_layer = QgsVectorLayer(orig_shp_path, src_layer.name(), "ogr")
        if not working_layer.isValid():
            self.log("Failed to load working layer.")
            return

        QgsProject.instance().addMapLayer(working_layer)
        self.log(f"Working layer loaded: {working_layer.name()} (for processing only)")


        # -------------------- Ensure PARCEL_ID --------------------
        field_names = [f.name() for f in working_layer.fields()]
        if 'PARCEL_ID' not in field_names:
            prov = working_layer.dataProvider()
            prov.addAttributes([QgsField('PARCEL_ID', QVariant.Int)])
            working_layer.updateFields()
            # assign integers based on X/Y order
            feats = list(working_layer.getFeatures())
            feats.sort(key=lambda f: (f.geometry().centroid().asPoint().x(),
                                    f.geometry().centroid().asPoint().y()))
            for i, f in enumerate(feats, start=1):
                f['PARCEL_ID'] = i
                prov.changeAttributeValues({f.id(): {working_layer.fields().indexFromName('PARCEL_ID'): i}})
            self.log(f"Created PARCEL_ID and assigned integers 1..{len(feats)}")

        # # -------------------- Shapefile regeneration process --------------------
        # self.log(f"Starting shapefile regeneration using ABR='{abr_col}', Built='{built_val}', Open='{open_val}'...")
        # # [Insert here your polygon merge, assign built_to_open, dissolve, built-only layer code from previous script]
        # # Ensure all outputs saved in `other_shapes_path`
        # # Keep self.log(...) messages for all major steps

        # -------------------- Shapefile regeneration process (final, combined exports) --------------------
        self.log(f"Starting shapefile regeneration using ABR='{abr_col}', Built='{built_val}', Open='{open_val}'...")

        # Use working_layer as source (already loaded earlier)
        src = working_layer
        fields = {f.name(): f for f in src.fields()}

        # Split built/open features
        built_feats = []
        open_feats = []
        for f in src.getFeatures():
            val = f[abr_col]
            if val is None:
                continue
            sval = str(val).strip().upper()
            if built_val.upper() in sval:
                built_feats.append(f)
            elif open_val.upper() in sval:
                open_feats.append(f)

        self.log(f"Found {len(built_feats)} BUILT_UP and {len(open_feats)} OPEN_SP features.")

        # Build spatial index for open features
        open_index = QgsSpatialIndex()
        open_id_to_feat = {}
        for of in open_feats:
            open_index.insertFeature(of)
            open_id_to_feat[of.id()] = of

        # Helper geometry metric functions (unchanged)
        VERT_TOL = 1e-6
        def vertex_set(geom, tol=VERT_TOL):
            pts = set()
            if geom is None or geom.isEmpty():
                return pts
            it = geom.vertices()
            while True:
                try:
                    p = next(it)
                    if tol > 0:
                        decimals = max(0, int(abs(math.log10(tol))))
                        pts.add((round(p.x(), decimals), round(p.y(), decimals)))
                    else:
                        pts.add((p.x(), p.y()))
                except StopIteration:
                    break
            return pts

        def shared_vertex_count(geomA, geomB, tol=VERT_TOL):
            return len(vertex_set(geomA, tol).intersection(vertex_set(geomB, tol)))

        def shared_boundary_length(geomA, geomB):
            if geomA is None or geomB is None:
                return 0.0
            try:
                inter = geomA.boundary().intersection(geomB.boundary())
                if inter is None or inter.isEmpty():
                    return 0.0
                return inter.length()
            except Exception:
                return 0.0

        # Build candidate metrics for merging (unchanged)
        built_candidates = {}
        for bf in built_feats:
            bid = bf.id()
            geom_b = bf.geometry()
            if geom_b is None or geom_b.isEmpty():
                built_candidates[bid] = []
                continue
            bbox = geom_b.boundingBox()
            candidate_open_ids = open_index.intersects(bbox)
            clist = []
            for ofid in candidate_open_ids:
                of = open_id_to_feat.get(ofid)
                if of is None:
                    continue
                geom_o = of.geometry()
                if geom_o is None or geom_o.isEmpty():
                    continue
                if not (geom_b.intersects(geom_o) or geom_b.touches(geom_o) or geom_b.within(geom_o) or geom_b.overlaps(geom_o)):
                    continue
                within_flag = geom_b.within(geom_o)
                sv = shared_vertex_count(geom_b, geom_o, tol=VERT_TOL)
                sl = shared_boundary_length(geom_b, geom_o)
                try:
                    overlap_area = geom_b.intersection(geom_o).area()
                except Exception:
                    overlap_area = 0.0
                open_pid = of['PARCEL_ID'] if 'PARCEL_ID' in of.fields().names() else None
                clist.append((ofid, within_flag, sv, sl, overlap_area, open_pid))
            built_candidates[bid] = clist

        # Assign built to best open (unchanged)
        built_to_open = {}
        open_to_built_list = {}
        for bf in built_feats:
            bid = bf.id()
            candidates = built_candidates.get(bid, [])
            qualifying = [c for c in candidates if (c[1] is True) or (c[2] >= 3)]
            if not qualifying:
                continue
            qualifying.sort(key=lambda c: (int(c[1]), c[2], c[3], c[4], -float(c[5]) if c[5] is not None else 0.0), reverse=True)
            chosen_open_fid = qualifying[0][0]
            built_to_open[bid] = chosen_open_fid
            open_to_built_list.setdefault(chosen_open_fid, []).append(bid)

        self.log(f"Assigned {len(built_to_open)} BUILT_UP features to OPEN parcels.")

        # Create memory layer for merged polygons (temporary)
        out_fields = src.fields().toList()
        if 'PARCEL_ID_final' not in [f.name() for f in out_fields]:
            out_fields.append(QgsField('PARCEL_ID_final', QVariant.String))

        mem_uri = f"Polygon?crs={src.crs().authid()}&index=yes"
        out_layer = QgsVectorLayer(mem_uri, f"{village_folder_name}_merged_temp", "memory")
        prov = out_layer.dataProvider()
        prov.addAttributes(out_fields)
        out_layer.updateFields()

        built_id_to_feat = {f.id(): f for f in built_feats}
        open_id_to_feat = {f.id(): f for f in open_feats}

        def copy_attrs(src_feat, dest_feat):
            for fld in src.fields():
                if fld.name() in dest_feat.fields().names():
                    dest_feat[fld.name()] = src_feat[fld.name()]

        used_open_ids = set()
        used_built_ids = set()

        # 1) For each OPEN that has assigned built(s): union the open with all its built polygons
        for open_fid, built_list in open_to_built_list.items():
            of = open_id_to_feat.get(open_fid)
            if of is None:
                continue
            geom_union = of.geometry()
            used_open_ids.add(open_fid)
            for bfid in built_list:
                bf = built_id_to_feat.get(bfid)
                if bf is None:
                    continue
                used_built_ids.add(bfid)
                try:
                    geom_union = geom_union.combine(bf.geometry())
                except Exception:
                    try:
                        geom_union = geom_union.union(bf.geometry())
                    except Exception:
                        pass
            newf = QgsFeature(out_layer.fields())
            newf.setGeometry(geom_union)
            copy_attrs(of, newf)
            newf['PARCEL_ID_final'] = str(of['PARCEL_ID'])
            prov.addFeature(newf)

        # 2) Add BUILT features that were NOT assigned (remain independent)
        for bf in built_feats:
            if bf.id() in used_built_ids:
                continue
            newf = QgsFeature(out_layer.fields())
            newf.setGeometry(bf.geometry())
            copy_attrs(bf, newf)
            newf['PARCEL_ID_final'] = str(bf['PARCEL_ID'])
            prov.addFeature(newf)

        # 3) Add OPEN features that were NOT used/assigned (remain independent)
        for of in open_feats:
            if of.id() in used_open_ids:
                continue
            newf = QgsFeature(out_layer.fields())
            newf.setGeometry(of.geometry())
            copy_attrs(of, newf)
            newf['PARCEL_ID_final'] = str(of['PARCEL_ID'])
            prov.addFeature(newf)

        out_layer.updateExtents()

        # === Save merged temporary layer as GPKG (Other Shape files) and load it ===
        merged_gpkg = os.path.join(other_shapes_path, f"{village_folder_name}_merged.gpkg")
        save_opts_tmp = QgsVectorFileWriter.SaveVectorOptions()
        save_opts_tmp.driverName = "GPKG"
        save_opts_tmp.layerName = f"{village_folder_name}_merged"

        errm = QgsVectorFileWriter.writeAsVectorFormatV2(
            out_layer, merged_gpkg, QgsProject.instance().transformContext(), save_opts_tmp
        )
        if errm[0] != QgsVectorFileWriter.NoError:
            self.log(f"Error saving merged gpkg: {errm}")
        else:
            self.log(f"Saved merged intermediate (gpkg): {merged_gpkg}")

        # Load merged gpkg into project (since you chose Option B)
        merged_loaded = QgsVectorLayer(f"{merged_gpkg}|layername={save_opts_tmp.layerName}", f"{village_folder_name}_merged", "ogr")
        if merged_loaded.isValid():
            QgsProject.instance().addMapLayer(merged_loaded)
            self.log("Loaded merged GPKG into project (temporary).")
        else:
            self.log("Warning: merged GPKG failed to load; will try to use memory out_layer for dissolve.")

        # Dissolve by PARCEL_ID_final (use processing) — result in memory
        diss_input = merged_loaded if merged_loaded.isValid() else out_layer
        params = {'INPUT': diss_input, 'FIELD': ['PARCEL_ID_final'], 'OUTPUT': 'memory:'}
        diss = processing.run("native:dissolve", params)
        dissolved = diss['OUTPUT']

        # === Save dissolved result as GPKG (Other Shape files) and load it ===
        parcel_gpkg = os.path.join(other_shapes_path, "PARCEL_ID_final.gpkg")
        save_opts_parcel = QgsVectorFileWriter.SaveVectorOptions()
        save_opts_parcel.driverName = "GPKG"
        save_opts_parcel.layerName = "PARCEL_ID_final"

        errp = QgsVectorFileWriter.writeAsVectorFormatV2(
            dissolved, parcel_gpkg, QgsProject.instance().transformContext(), save_opts_parcel
        )
        if errp[0] != QgsVectorFileWriter.NoError:
            self.log(f"Error saving parcel gpkg: {errp}")
        else:
            self.log(f"Saved parcel (gpkg): {parcel_gpkg}")

        parcel_loaded = QgsVectorLayer(f"{parcel_gpkg}|layername={save_opts_parcel.layerName}", "Parcel Area (GPKG)", "ogr")
        if parcel_loaded.isValid():
            QgsProject.instance().addMapLayer(parcel_loaded)
            self.log("Loaded PARCEL_ID_final GPKG into project.")
        else:
            self.log("Warning: PARCEL_ID_final GPKG failed to load.")

        # === Create Built-only memory layer and save as GPKG (Other Shape files) and load it ===
        mem_uri_built = f"Polygon?crs={src.crs().authid()}&index=yes"
        built_layer = QgsVectorLayer(mem_uri_built, "Built up Area_temp", "memory")
        prov_built = built_layer.dataProvider()
        prov_built.addAttributes(src.fields())
        built_layer.updateFields()

        count_built = 0
        for f in src.getFeatures():
            abr = f[abr_col]
            if abr is None:
                continue
            if built_val.upper() in str(abr).strip().upper():
                newf = QgsFeature(built_layer.fields())
                newf.setGeometry(f.geometry())
                copy_attrs(f, newf)
                prov_built.addFeature(newf)
                count_built += 1

        built_layer.updateExtents()
        self.log(f"Built-only memory layer created with {count_built} features.")

        built_gpkg = os.path.join(other_shapes_path, "Built_only.gpkg")
        save_opts_built = QgsVectorFileWriter.SaveVectorOptions()
        save_opts_built.driverName = "GPKG"
        save_opts_built.layerName = "Built_only"

        errb = QgsVectorFileWriter.writeAsVectorFormatV2(
            built_layer, built_gpkg, QgsProject.instance().transformContext(), save_opts_built
        )
        if errb[0] != QgsVectorFileWriter.NoError:
            self.log(f"Error saving built-only gpkg: {errb}")
        else:
            self.log(f"Saved built-only (gpkg): {built_gpkg}")

        built_loaded = QgsVectorLayer(f"{built_gpkg}|layername={save_opts_built.layerName}", "Built up Area (GPKG)", "ogr")
        if built_loaded.isValid():
            QgsProject.instance().addMapLayer(built_loaded)
            self.log("Loaded Built_only GPKG into project.")
        else:
            self.log("Warning: Built_only GPKG failed to load.")

        # -------------------- FINAL EXPORTS (write SHP outputs into Shapefiles (Reprocessed)) --------------------
        # Parcel Area.shp (from dissolved)
        parcel_shp = os.path.join(reprocessed_path, "Parcel Area.shp")
        save_opts_shp = QgsVectorFileWriter.SaveVectorOptions()
        save_opts_shp.driverName = "ESRI Shapefile"
        save_opts_shp.layerName = "Parcel Area"

        err_shp_p = QgsVectorFileWriter.writeAsVectorFormatV2(
            dissolved, parcel_shp, QgsProject.instance().transformContext(), save_opts_shp
        )
        if err_shp_p[0] != QgsVectorFileWriter.NoError:
            self.log(f"Error saving Parcel Area.shp: {err_shp_p}")
        else:
            self.log(f"Saved Parcel Area.shp: {parcel_shp}")

        # Built up Area.shp (from built_layer)
        built_shp = os.path.join(reprocessed_path, "Built up Area.shp")
        save_opts_shp2 = QgsVectorFileWriter.SaveVectorOptions()
        save_opts_shp2.driverName = "ESRI Shapefile"
        save_opts_shp2.layerName = "Built up Area"

        err_shp_b = QgsVectorFileWriter.writeAsVectorFormatV2(
            built_layer, built_shp, QgsProject.instance().transformContext(), save_opts_shp2
        )
        if err_shp_b[0] != QgsVectorFileWriter.NoError:
            self.log(f"Error saving Built up Area.shp: {err_shp_b}")
        else:
            self.log(f"Saved Built up Area.shp: {built_shp}")

        # -------------------- SAVE TEMP LAYERS (GPKG) & KEEP THEM LOADED (Option B) --------------------
        # (The gpkg saves were already created above; ensure they exist and are logged)
        self.log("Temporary GPKG files saved to Other Shape files and currently loaded in project (if valid).")

        # -------------------- CLEANUP memory-only temp layers (remove *_temp memory layers) --------------------
        self.log("Cleaning up in-memory temporary layers (keeping GPKG-backed layers)...")

        mem_remove_keys = ["_temp"]
        for lyr in list(QgsProject.instance().mapLayers().values()):
            if any(k in lyr.name() for k in mem_remove_keys) and lyr.providerType() == 'memory':
                try:
                    QgsProject.instance().removeMapLayer(lyr.id())
                    self.log(f"Removed memory layer: {lyr.name()}")
                except Exception:
                    pass

        # -------------------- Ensure final SHPs are loaded (load them if not already) --------------------
        loaded_layers = []
        if os.path.exists(parcel_shp):
            pl = QgsVectorLayer(parcel_shp, "Parcel Area", "ogr")
            if pl.isValid():
                QgsProject.instance().addMapLayer(pl)
                loaded_layers.append("Parcel Area")
                self.log("Loaded Parcel Area.shp into project.")
            else:
                self.log("Failed to load Parcel Area.shp")
        else:
            self.log("Parcel Area.shp not found on disk.")

        if os.path.exists(built_shp):
            bl = QgsVectorLayer(built_shp, "Built up Area", "ogr")
            if bl.isValid():
                QgsProject.instance().addMapLayer(bl)
                loaded_layers.append("Built up Area")
                self.log("Loaded Built up Area.shp into project.")
            else:
                self.log("Failed to load Built up Area.shp")
        else:
            self.log("Built up Area.shp not found on disk.")

        # -------------------- Set visibility: only Built up Area & Parcel Area --------------------
        project = QgsProject.instance()
        root = project.layerTreeRoot()

        parcel_layer_name = "Parcel Area"
        built_layer_name = "Built up Area"

        parcel_node = None
        built_node = None

        # First pass: hide all layers & find nodes
        for node in root.children():
            if isinstance(node, QgsLayerTreeLayer):
                # Hide all layers
                node.setItemVisibilityChecked(False)

                lyr = node.layer()
                if lyr is None:
                    continue

                # Match by layer name
                if lyr.name() == parcel_layer_name:
                    parcel_node = node
                elif lyr.name() == built_layer_name:
                    built_node = node

        # Second pass: enable the final two layers
        if parcel_node:
            parcel_node.setItemVisibilityChecked(True)
        else:
            self.log("Warning: Parcel Area layer not found.")

        if built_node:
            built_node.setItemVisibilityChecked(True)
        else:
            self.log("Warning: Built up Area layer not found.")

        # OPTIONAL: Move both layers to the TOP of the layer panel
        # (QGIS equivalent to Ctrl+Shift+H for visibility cleanup)

        if parcel_node:
            cloned = parcel_node.clone()
            root.insertChildNode(0, cloned)
            root.removeChildNode(parcel_node)
            
        if built_node:
            cloned = built_node.clone()
            root.insertChildNode(0, cloned)
            root.removeChildNode(built_node)


        self.log("Layer visibility updated: Only 'Parcel Area' and 'Built up Area' are visible.")


        # -------------------- Final Project Save --------------------
        final_proj_path = os.path.join(backup_path, f"{village_folder_name} (Reprocessed).qgz")
        QgsProject.instance().write(final_proj_path)
        
        self.log(f"Final project saved: {final_proj_path}")

        self.log("Shapefile regeneration completed successfully.")

        self.log("Processing finished.")
