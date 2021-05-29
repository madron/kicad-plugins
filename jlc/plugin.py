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
        self.layers = [
            dict(suffix='CuTop',        format=pcbnew.F_Cu,         description='Top layer'),
            dict(suffix='CuBottom',     format=pcbnew.B_Cu,         description='Bottom layer'),
            dict(suffix='PasteTop',     format=pcbnew.F_Paste,      description='Paste top'),
            dict(suffix='PasteBottom',  format=pcbnew.B_Paste,      description='Paste Bottom'),
            dict(suffix='SilkTop',      format=pcbnew.F_SilkS,      description='Silk top'),
            dict(suffix='SilkBottom',   format=pcbnew.B_SilkS,      description='Silk top'),
            dict(suffix='MaskTop',      format=pcbnew.F_Mask,       description='Mask top'),
            dict(suffix='MaskBottom',   format=pcbnew.B_Mask,       description='Mask bottom'),
            dict(suffix='EdgeCuts',     format=pcbnew.Edge_Cuts,    description='Edges'),
        ]
        for layer in self.layers:
            layer['file_name'] = '{}.gbr'.format(layer['suffix'])
            layer['file_path'] = os.path.join(self.fab_dir, layer['file_name'])
        self.drill_name = 'Drill.drl'
        self.drill_path = os.path.join(self.fab_dir, self.drill_name)
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

        popt.SetOutputDirectory(self.fab_dir)

        # Set some important plot options:
        # One cannot plot the frame references, because the board does not know
        # the frame references.
        popt.SetPlotFrameRef(False)
        # popt.SetSketchPadLineWidth(pcbnew.FromMM(0.1))

        popt.SetAutoScale(False)
        popt.SetScale(1)
        popt.SetMirror(False)
        popt.SetUseGerberAttributes(True)
        popt.SetExcludeEdgeLayer(False);
        popt.SetScale(1)
        popt.SetUseAuxOrigin(True)

        # This by gerbers only (also the name is truly horrid!)
        popt.SetSubtractMaskFromSilk(False) #remove solder mask from silk to be sure there is no silk on pads

        for layer in self.layers:
            pctl.SetLayer(layer['format'])
            pctl.OpenPlotfile(layer['suffix'], pcbnew.PLOT_FORMAT_GERBER, layer['description'])
            pctl.PlotLayer()

        pctl.ClosePlot()

        for layer in self.layers:
            source_name = os.path.join(self.fab_dir, '{}-{}'.format(self.project_name, layer['file_name']))
            os.rename(source_name, layer['file_path'])

    def generate_drill(self):
        # Fabricators need drill files.
        # sometimes a drill map file is asked (for verification purpose)
        drill_writer = pcbnew.EXCELLON_WRITER(self.board)
        drill_writer.SetMapFileFormat(pcbnew.PLOT_FORMAT_PDF)

        mirror = False
        minimal_header = False
        offset = pcbnew.wxPoint(0,0)
        # False to generate 2 separate drill files (one for plated holes, one for non plated holes)
        # True to generate only one drill file
        merge_npth = True
        drill_writer.SetOptions(mirror, minimal_header, offset, merge_npth)

        metric_format = True
        drill_writer.SetFormat(metric_format)

        generate_drill = True
        generate_map = False
        drill_writer.CreateDrillandMapFilesSet(self.fab_dir, generate_drill, generate_map)

        source_path = os.path.join(self.fab_dir, '{}.drl'.format(self.project_name))
        os.rename(source_path, self.drill_path)

    def generate_gerber_zipfile(self):
        gerber = ZipFile(os.path.join(self.jlc_dir, '{}.zip'.format(self.project_name)), 'w')
        for layer in self.layers:
            gerber.write(layer['file_path'], layer['file_name'])
        gerber.write(self.drill_path, self.drill_name)
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
            out.writerow(['Comment', 'Designator', 'Footprint', 'LCSC Part #'])
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
        self.generate_gerber_zipfile()
        self.generate_bom()
        self.generate_position()


# register plugin with kicad backend
JlcPlugin().register()
