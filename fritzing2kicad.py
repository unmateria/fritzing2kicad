import sys
import os
import zipfile
import re
import math
import statistics
import xml.etree.ElementTree as ET
from io import BytesIO
from svgelements import SVG, Circle, Rect, Path, Group

class FritzingConverter:
    def __init__(self):
        self.scale_to_mm = 1.0
        self.min_x = float('inf')
        self.min_y = float('inf')
        self.max_x = float('-inf')
        self.max_y = float('-inf')

    def find_file_in_zip(self, z, partial_name):
        if not partial_name:
            candidates = [f for f in z.namelist() if f.endswith('.svg') and 'pcb' in f.lower()]
            if candidates: return candidates[0]
            raise ValueError("Image path is empty and no alternative found.")
            
        target = os.path.basename(partial_name.replace('\\', '/'))
        candidates = [f for f in z.namelist() if f.endswith(target)]
        
        if len(candidates) > 1:
            if 'pcb' in target and 'pcb' in partial_name:
                best = [c for c in candidates if 'pcb' in c.lower() and 'schematic' not in c.lower()]
                if best: return best[0]
        
        if candidates:
            return candidates[0]
        
        fallback = [f for f in z.namelist() if 'pcb' in f.lower() and f.endswith('.svg')]
        if fallback: return fallback[0]
        
        raise KeyError(f"Could not find {target} in ZIP")

    def parse_metadata(self, fzp_xml):
        root = ET.fromstring(fzp_xml)
        ns = {'f': 'http://fritzing.org/dtd/2.0'}
        
        title = root.find('.//f:title', ns)
        if title is None: title = root.find('.//title')
        comp_name = title.text if title is not None else "Component"
        comp_name = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', comp_name)

        pcb_view = root.find('.//f:pcbView', ns) or root.find('.//pcbView')
        pcb_image = None
        if pcb_view is not None:
            ly = pcb_view.find('.//f:layers', ns) or pcb_view.find('.//layers')
            if ly is not None: pcb_image = ly.get('image')

        connectors = []
        clist = root.find('.//f:connectors', ns) or root.find('.//connectors')
        if clist is not None:
            for conn in clist:
                cid = conn.get('id')
                
                desc_tag = conn.find('.//f:description', ns)
                if desc_tag is None: desc_tag = conn.find('.//description')
                
                name_tag = conn.get('name')
                
                final_name = name_tag
                if desc_tag is not None and desc_tag.text:
                    clean_desc = desc_tag.text.strip()
                    if len(clean_desc) > 0:
                        final_name = clean_desc

                if len(final_name) > 15: final_name = final_name[:15]
                if final_name.lower() == "passive": final_name = cid 

                pin_num = cid
                if cid.startswith('connector'):
                    try:
                        pin_num = str(int(cid.replace('connector', '')) + 1)
                    except: pass
                
                svg_id = None
                is_tht = False
                pview = conn.find('.//f:pcbView', ns) or conn.find('.//pcbView')
                if pview is not None:
                    has_c0 = False
                    has_c1 = False
                    for p in (pview.findall('.//f:p', ns) or pview.findall('.//p')):
                        lid = p.get('layer')
                        if lid == 'copper0': has_c0 = True
                        if lid == 'copper1': has_c1 = True
                        svg_id = p.get('svgId')
                    is_tht = (has_c0 and has_c1)

                connectors.append({
                    'id': cid,
                    'name': final_name,
                    'pin': pin_num,
                    'svg_id': svg_id,
                    'is_tht': is_tht
                })
        
        try:
            connectors.sort(key=lambda x: int(x['pin']))
        except: pass 

        return comp_name, pcb_image, connectors

    def calculate_scale_from_pitch(self, svg_content, connectors):
        try:
            svg_obj = SVG.parse(BytesIO(svg_content))
            conn_ids = {c['svg_id'] for c in connectors if c['svg_id']}
            y_coords = []
            
            def find_centroids(element):
                if isinstance(element, Group):
                    for child in element: find_centroids(child)
                    return
                if element.id in conn_ids:
                    try:
                        bbox = element.bbox()
                        if bbox:
                            cy = (bbox[1] + bbox[3]) / 2.0
                            y_coords.append(cy)
                    except: pass
            
            for elem in svg_obj.elements():
                find_centroids(elem)
            
            if len(y_coords) < 2:
                print("WARNING: Too few pins to auto-calibrate. Using fallback 1 mil.")
                return 0.0254 

            y_coords.sort()
            deltas = []
            for i in range(1, len(y_coords)):
                d = y_coords[i] - y_coords[i-1]
                if d > 0.1:
                    deltas.append(d)

            if not deltas: return 0.0254

            rounded_deltas = [round(d, 1) for d in deltas]
            mode_delta = statistics.mode(rounded_deltas)
            
            scale = 2.54 / mode_delta
            
            print(f"DEBUG: Auto-calibration -> SVG Distance: {mode_delta:.2f} units = 2.54mm. Factor: {scale:.6f}")
            return scale

        except Exception as e:
            print(f"Calibration error: {e}. Using fallback.")
            return 0.0254

    def generate_footprint(self, z, pcb_image_path, connectors, comp_name, output_file):
        try:
            real_path = self.find_file_in_zip(z, pcb_image_path)
            svg_content = z.read(real_path)
        except (KeyError, ValueError) as e:
            print(f"WARNING: PCB SVG not found. Skipping footprint.")
            return

        self.scale_to_mm = self.calculate_scale_from_pitch(svg_content, connectors)
        
        self.min_x, self.min_y = float('inf'), float('inf')
        self.max_x, self.max_y = float('-inf'), float('-inf')

        svg_obj = SVG.parse(BytesIO(svg_content))
        conn_map = {c['svg_id']: c for c in connectors if c['svg_id']}
        pads = []

        def process_element(element, inherited_id=None):
            if isinstance(element, Group):
                for child in element:
                    process_element(child, inherited_id or element.id)
                return

            eid = element.id or inherited_id
            if eid in conn_map:
                data = conn_map[eid]
                try:
                    bbox = element.bbox()
                    if not bbox: return
                except: return

                cx = (bbox[0] + bbox[2]) / 2.0 * self.scale_to_mm
                cy = (bbox[1] + bbox[3]) / 2.0 * self.scale_to_mm
                w = (bbox[2] - bbox[0]) * self.scale_to_mm
                h = (bbox[3] - bbox[1]) * self.scale_to_mm

                self.min_x = min(self.min_x, cx)
                self.min_y = min(self.min_y, cy)
                self.max_x = max(self.max_x, cx)
                self.max_y = max(self.max_y, cy)

                pad_type = "thru_hole" if data['is_tht'] else "smd"
                shape = "rect"
                drill = 0
                layers = "*.Cu *.Mask" if data['is_tht'] else "F.Cu F.Mask F.Paste"

                if isinstance(element, Circle):
                    shape = "circle"
                    calc_drill = min(w, h) * 0.65
                    if data['is_tht']: drill = max(0.6, calc_drill)
                elif isinstance(element, Rect):
                    shape = "rect"
                    if data['is_tht']: drill = min(w, h) * 0.6
                elif isinstance(element, Path):
                    shape = "custom"
                
                pad_def = {
                    'pin': data['pin'], 'type': pad_type, 'shape': shape,
                    'pos': (cx, cy), 'size': (w, h), 'layers': layers, 'drill': drill
                }
                if shape == "custom":
                    pts = []
                    try:
                        for seg in element.as_points(n=6):
                            pts.append((seg.start.x * self.scale_to_mm, seg.start.y * self.scale_to_mm))
                    except: pass
                    pad_def['points'] = pts
                
                pads.append(pad_def)

        for elem in svg_obj.elements():
            process_element(elem)

        if self.min_x == float('inf'):
            print("WARNING: No pads detected.")
            return

        cx_total = (self.min_x + self.max_x) / 2.0
        cy_total = (self.min_y + self.max_y) / 2.0
        h_total = self.max_y - self.min_y
        
        with open(output_file, 'w') as f:
            f.write(f'(footprint "{comp_name}" (layer "F.Cu")\n')
            f.write('  (attr smd)\n')
            f.write(f'  (fp_text reference "REF**" (at 0 -{h_total/2 + 2:.2f}) (layer "F.SilkS") (effects (font (size 1 1) (thickness 0.15))))\n')
            f.write(f'  (fp_text value "{comp_name}" (at 0 {h_total/2 + 2:.2f}) (layer "F.Fab") (effects (font (size 1 1) (thickness 0.15))))\n')
            
            for p in pads:
                pos_x = p['pos'][0] - cx_total
                pos_y = p['pos'][1] - cy_total
                at = f"(at {pos_x:.4f} {pos_y:.4f})"
                sz = f"(size {p['size'][0]:.4f} {p['size'][1]:.4f})"
                
                if p['shape'] == 'custom' and p.get('points'):
                    pts_str = " ".join([f"(xy {pt[0]-cx_total:.4f} {pt[1]-cy_total:.4f})" for pt in p['points']])
                    f.write(f'  (pad "{p["pin"]}" {p["type"]} custom {at} {sz} (layers "{p["layers"]}")')
                    if p['drill']: f.write(f' (drill {p["drill"]:.4f})')
                    f.write('\n    (options (clearance outline) (anchor circle))\n')
                    f.write(f'    (primitives (gr_poly (pts {pts_str}) (width 0))))\n')
                else:
                    drill_s = f' (drill {p["drill"]:.4f})' if p['drill'] else ''
                    f.write(f'  (pad "{p["pin"]}" {p["type"]} {p["shape"]} {at} {sz} (layers "{p["layers"]}"){drill_s})\n')
            f.write(')')
        print(f"Footprint generated: {output_file}")

    def generate_symbol(self, connectors, comp_name, output_file):
        count = len(connectors)
        half = math.ceil(count / 2)
        pin_spacing = 2.54 
        height = max(half * pin_spacing, 5.08) + 2.54
        max_name_len = max([len(c['name']) for c in connectors]) if connectors else 5
        width = max(15.24, max_name_len * 1.5)
        
        with open(output_file, 'w') as f:
            f.write('(kicad_symbol_lib (version 20211014) (generator "Fritzing2KiCad")\n')
            f.write(f'  (symbol "{comp_name}" (in_bom yes) (on_board yes)\n')
            f.write(f'    (property "Reference" "U" (id 0) (at -5.08 {height/2 + 2.54} 0) (effects (font (size 1.27 1.27))))\n')
            f.write(f'    (property "Value" "{comp_name}" (id 1) (at 0 {height/2 + 2.54} 0) (effects (font (size 1.27 1.27))))\n')
            f.write(f'    (property "Footprint" "" (id 2) (at 0 -{height/2 + 2.54} 0) (effects (font (size 1.27 1.27)) (hide yes)))\n')
            f.write(f'    (symbol "{comp_name}_1_1"\n')
            f.write(f'      (rectangle (start -{width/2:.2f} {height/2:.2f}) (end {width/2:.2f} -{height/2:.2f}) (stroke (width 0.254)) (fill (type background)))\n')
            f.write('    )\n') 
            f.write(f'    (symbol "{comp_name}_1_1"\n')
            
            y_pos = height / 2 - 1.27
            for i, conn in enumerate(connectors):
                p_name = conn['name']
                p_num = conn['pin']
                
                pin_type = "passive" 

                if i < half:
                    x = -width/2 - 2.54
                    y = y_pos - (i * pin_spacing)
                    f.write(f'      (pin {pin_type} line (at {x:.2f} {y:.2f} 0) (length 2.54)\n')
                    f.write(f'        (name "{p_name}" (effects (font (size 1.27 1.27))))\n')
                    f.write(f'        (number "{p_num}" (effects (font (size 1.27 1.27))))\n')
                    f.write('      )\n')
                else:
                    idx_right = i - half
                    x = width/2 + 2.54
                    y = y_pos - (idx_right * pin_spacing)
                    f.write(f'      (pin {pin_type} line (at {x:.2f} {y:.2f} 180) (length 2.54)\n')
                    f.write(f'        (name "{p_name}" (effects (font (size 1.27 1.27))))\n')
                    f.write(f'        (number "{p_num}" (effects (font (size 1.27 1.27))))\n')
                    f.write('      )\n')

            f.write('    )\n') 
            f.write('  )\n') 
            f.write(')\n') 
            print(f"Symbol generated: {output_file}")

    def process(self, fzpz_file, out_base_name):
        if not fzpz_file.endswith('.fzpz'):
             print("Error: Input file must be .fzpz")
             return

        with zipfile.ZipFile(fzpz_file, 'r') as z:
            try:
                fzp_name = [f for f in z.namelist() if f.endswith('.fzp')][0]
            except IndexError:
                print("Error: No .fzp file found in archive.")
                return
            
            fzp_xml = z.read(fzp_name)
            comp_name_meta, pcb_image, connectors = self.parse_metadata(fzp_xml)
            
            final_name = os.path.basename(out_base_name).split('.')[0] if out_base_name else comp_name_meta
            final_name = re.sub(r'[^a-zA-Z0-9_\-]', '_', final_name)

            print(f"Processing: '{final_name}' ({len(connectors)} pins)")
            
            mod_file = out_base_name + ".kicad_mod"
            self.generate_footprint(z, pcb_image, connectors, final_name, mod_file)
            
            sym_file = out_base_name + ".kicad_sym"
            self.generate_symbol(connectors, final_name, sym_file)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python fritzing2kicad.py input.fzpz output_name")
    else:
        conv = FritzingConverter()
        conv.process(sys.argv[1], sys.argv[2])
