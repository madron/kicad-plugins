import csv
import os
import shutil
import sys
import pcbnew
import yaml
from collections import OrderedDict
from zipfile import ZipFile


class JlcPlugin(pcbnew.ActionPlugin):
    def defaults(self):
        self.plugin_dir = os.path.dirname(__file__)
        self.name = 'Jlc'
        self.description = 'Jlc pcb fabrication and smt assembly'
        self.category = 'Pcb fabrication'
        self.show_toolbar_button = True
        self.icon_file_name = os.path.join(self.plugin_dir, 'icon.png')

    def prepare(self):
        self.board: pcbnew.BOARD = pcbnew.GetBoard()
        self.project_dir = os.path.dirname(self.board.GetFileName())
        self.project_name = os.path.splitext(os.path.basename(self.board.GetFileName()))[0]
        self.fab_dir = os.path.join(self.project_dir, 'fab')
        self.jlc_dir = os.path.join(self.project_dir, 'jlc')
        self.files = [
            dict(source_extension='-drl_map.gbr',   name='Drill_Map.gbr'),
            dict(source_extension='.drl',           name='Drill.drl'),
            # dict(source_extension='-PTH.drl',           name='PTH.drl'),
            # dict(source_extension='-NPTH.drl',          name='NPTH.drl'),
            # dict(source_extension='-PTH-drl_map.gbr',   name='PTH-drl_map.drl'),
            # dict(source_extension='-NPTH-drl_map.gbr',  name='NPTH-drl_map.drl'),
        ]
        self.layers = [
            dict(suffix='F_Cu',      extension='gtl', format=pcbnew.F_Cu,      description='Top layer'),
            dict(suffix='B_Cu',      extension='gbl', format=pcbnew.B_Cu,      description='Bottom layer'),
            dict(suffix='F_Paste',   extension='gtp', format=pcbnew.F_Paste,   description='Paste top'),
            # dict(suffix='B_Paste',   extension='gbp', format=pcbnew.B_Paste,   description='Paste Bottom'),
            dict(suffix='F_SilkS',   extension='gto', format=pcbnew.F_SilkS,   description='Silk top'),
            dict(suffix='B_SilkS',   extension='gbo', format=pcbnew.B_SilkS,   description='Silk top'),
            dict(suffix='F_Mask',    extension='gts', format=pcbnew.F_Mask,    description='Mask top'),
            dict(suffix='B_Mask',    extension='gbs', format=pcbnew.B_Mask,    description='Mask bottom'),
            dict(suffix='Edge_Cuts', extension='gm1', format=pcbnew.Edge_Cuts, description='Edges'),
        ]
        for layer in self.layers:
            layer['file_name'] = '{}.{}'.format(layer['suffix'], layer['extension'])
            layer['file_path'] = os.path.join(self.fab_dir, layer['file_name'])
            name = '{}.{}'.format(layer['suffix'], layer['extension'])
            rename = dict(
                source_extension='-{}'.format(name),
                name=name,
            )
            self.files.append(rename)
        for rename_file in self.files:
            source_name = '{}{}'.format(self.project_name, rename_file['source_extension'])
            rename_file['source_path'] = os.path.join(self.fab_dir, source_name)
            rename_file['destination_path'] = os.path.join(self.fab_dir, rename_file['name'])
        self.drill_names = ['NPTH.drl', 'PTH.drl', 'drl_map.gbr']

        self.bom_name = 'bom.csv'
        self.bom_path = os.path.join(self.jlc_dir, self.bom_name)
        self.rotation_override_name = 'jlc-rotation-override.yml'
        self.rotation_override_path = os.path.join(self.project_dir, self.rotation_override_name)
        self.position_name = 'cpl.csv'
        self.position_path = os.path.join(self.jlc_dir, self.position_name)
        # cleanup fab_dir
        try:
            shutil.rmtree(self.fab_dir)
        except FileNotFoundError:
            pass
        os.mkdir(self.fab_dir)
        # cleanup jlc_dir
        try:
            shutil.rmtree(self.jlc_dir)
        except FileNotFoundError:
            pass
        os.mkdir(self.jlc_dir)

    def generate_gerber(self):
        pctl: pcbnew.PLOT_CONTROLLER = pcbnew.PLOT_CONTROLLER(self.board)
        popt: pcbnew.PCB_PLOT_PARAMS = pctl.GetPlotOptions()

        popt.SetFormat(pcbnew.PLOT_FORMAT_GERBER)
        popt.SetOutputDirectory(self.fab_dir)

        ### General options
        # Plot border and title block (False)
        # ?
        # Plot footprint values (False)
        popt.SetPlotValue(False)
        # Plot footprint references (True)
        popt.SetPlotReference(True)
        # Force plotting of invisible values/refs (False)
        popt.SetPlotInvisibleText(False)
        # Exclude PCB edge layer from other layers (True)
        popt.SetExcludeEdgeLayer(True)
        # Exclude pads from silkscreen (True)
        popt.SetPlotPadsOnSilkLayer(False)
        # Do not tent vias (False)
        # ?
        # Use auxiliary axis as origin (False)
        popt.SetUseAuxOrigin(False)
        #
        # Drill Marks (None)
        popt.SetDrillMarksType(pcbnew.PCB_PLOT_PARAMS.NO_DRILL_SHAPE)
        # Scaling (1:1)
        popt.SetAutoScale(False)
        popt.SetScale(1)
        # Plot mode (Filled)
        popt.SetPlotMode(pcbnew.FILLED_SHAPE)
        # Default line width (5.905512 mills)
        # popt.SetLineWidth( int ??? )
        # Mirrored plot (False)
        popt.SetMirror(False)
        # Negative plot (False)
        popt.SetNegative(False)
        # Check zone fills before plotting (True)
        # ?

        ### Gerber options
        # Use Protel file extension (True)
        popt.SetUseGerberProtelExtensions(True)
        # Generate gerber job file (False)
        popt.SetCreateGerberJobFile(False)
        # Subtract soldermask from silkscreen (True)
        popt.SetSubtractMaskFromSilk(True)
        # Coordinate format (4.6, unit mm)
        # ?
        # Use extended X2 format (False)
        popt.SetUseGerberX2format(False)
        # Include netlist attributes (False)
        popt.SetIncludeGerberNetlistInfo(False)

        ### Other
        # Set some important plot options:
        # One cannot plot the frame references, because the board does not know
        # the frame references.
        # popt.SetPlotFrameRef(False)
        # popt.SetUseGerberAttributes(True)

        for layer in self.layers:
            pctl.SetLayer(layer['format'])
            pctl.OpenPlotfile(layer['suffix'], pcbnew.PLOT_FORMAT_GERBER, layer['description'])
            pctl.PlotLayer()

        pctl.ClosePlot()

    def generate_drill(self):
        drill_writer = pcbnew.EXCELLON_WRITER(self.board)
        drill_writer.SetMapFileFormat(pcbnew.PLOT_FORMAT_GERBER)

        mirror = False
        minimal_header = True
        offset = pcbnew.wxPoint(0,0)
        merge_npth = True
        drill_writer.SetOptions(mirror, minimal_header, offset, merge_npth)

        # SetFormat(EXCELLON_WRITER self, bool aMetric, GENDRILL_WRITER_BASE::ZEROS_FMT aZerosFmt, int aLeftDigits=0, int aRightDigits=0)
        metric_format = True
        zeros_format = pcbnew.GENDRILL_WRITER_BASE.DECIMAL_FORMAT
        left_digits = 0
        right_digits = 0
        drill_writer.SetFormat(metric_format, zeros_format, left_digits, right_digits)

        generate_drill = True
        generate_map = True
        drill_writer.CreateDrillandMapFilesSet(self.fab_dir, generate_drill, generate_map)

    def rename_files(self):
        for f in self.files:
            os.rename(f['source_path'], f['destination_path'])

    def generate_gerber_zipfile(self):
        gerber = ZipFile(os.path.join(self.jlc_dir, '{}.zip'.format(self.project_name)), 'w')
        for f in self.files:
            gerber.write(f['destination_path'], f['name'])
        gerber.close()

    def generate_bom(self):
        sys.path.append('/usr/share/kicad/plugins')
        import kicad_netlist_reader
        # parse netlist
        netlist_name = '{}.xml'.format(self.project_name)
        netlist_path = os.path.join(self.project_dir, netlist_name)
        net = kicad_netlist_reader.netlist(netlist_path)
        # write csv
        with open(self.bom_path, 'w', newline='') as f:
            out = csv.writer(f)
            # out.writerow(['Comment', 'Designator', 'Footprint', 'LCSC Part #'])
            for group in net.groupComponents():
                refs = []
                lcsc_pn = ''
                for component in group:
                    ref = component.getRef()
                    if ref == 'REF**':
                        continue
                    lcsc_pn = component.getField("LCSC") or lcsc_pn
                    if lcsc_pn.lower() == 'skip':
                        continue
                    refs.append(ref)
                    c = component
                if len(refs) == 0:
                    continue
                # Fill in the component groups common data
                comment = ''
                out.writerow([c.getValue() + " " + c.getDescription(), ",".join(refs), c.getFootprint().split(':')[1], lcsc_pn])
            f.close()

    def generate_position(self):
        rotation_overrides = dict()
        if os.path.isfile(self.rotation_override_path):
            with open(self.rotation_override_path, 'r') as stream:
                rotation_overrides = yaml.safe_load(stream)
        with open(self.position_path, 'w', newline='') as f:
            fieldnames = ['Designator', 'Mid X', 'Mid Y', 'Layer', 'Rotation']
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for module in self.board.GetModules():
                module: pcbnew.MODULE = module
                if not module.GetLayerName() == 'F.Cu':
                    continue
                layer = 'T'
                ref = module.GetReference()
                if ref == 'REF**':
                    continue
                center = module.GetCenter()
                x = '{}mm'.format(center.x / 1000000)
                y = '{}mm'.format(-center.y / 1000000)
                rotation_override = rotation_overrides.get(ref, 0)
                rotation = module.GetOrientationDegrees()
                rotation = (360 + int(rotation) + rotation_override) % 360
                writer.writerow({
                    'Designator': ref,
                    'Mid X': x,
                    'Mid Y': y,
                    'Layer': layer,
                    'Rotation': rotation,
                })

    def Run(self):
        self.prepare()
        self.generate_gerber()
        self.generate_drill()
        self.rename_files()
        self.generate_gerber_zipfile()
        self.generate_bom()
        self.generate_position()


# register plugin with kicad backend
JlcPlugin().register()
