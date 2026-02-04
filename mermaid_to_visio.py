"""
Mermaid to Visio Converter with Connection Points and Master Shapes
Parses Mermaid flowchart/graph syntax and generates Visio shapes on A4 sheet

Usage:
    python mermaid_to_visio.py --file diagram.txt [options]
    python mermaid_to_visio.py --clipboard [options]

Options:
    --layout [flow|hilbert]     Layout algorithm (default: flow)
    --horizontal N              Horizontal connection points (default: 5)
    --vertical N                Vertical connection points (default: 3)
"""

import re
import win32com.client
import pyperclip
import argparse
import sys
import os
from collections import defaultdict, deque
import math

# A4 dimensions in inches (Visio uses inches)
A4_WIDTH = 8.27
A4_HEIGHT = 11.69

# Margins and shape sizing
MARGIN = 0.5
SHAPE_WIDTH = 1.5
SHAPE_HEIGHT = 0.75
HORIZONTAL_SPACING = 0.5
VERTICAL_SPACING = 0.75

# Connection point defaults
DEFAULT_HORIZONTAL_CONNECTIONS = 5
DEFAULT_VERTICAL_CONNECTIONS = 3


class MermaidParser:
    """Parse Mermaid diagram syntax"""

    def __init__(self, mermaid_text):
        self.text = mermaid_text
        self.nodes = {}  # node_id -> label
        self.edges = []  # list of (from_id, to_id)

    def parse(self):
        """Extract nodes and edges from Mermaid syntax"""
        lines = self.text.strip().split('\n')

        for line in lines:
            line = line.strip()

            # Skip diagram type declarations and empty lines
            if not line or line.startswith('graph') or line.startswith('flowchart'):
                continue

            # Parse node definitions and connections
            # Patterns like: A[Label] or A --> B or A[Label] --> B[Label]

            # Connection pattern: A --> B or A[Label] --> B[Label]
            connection_match = re.search(r'(\w+)(?:\[([^\]]+)\])?\s*--[->]+\s*(\w+)(?:\[([^\]]+)\])?', line)
            if connection_match:
                from_id = connection_match.group(1)
                from_label = connection_match.group(2) or from_id
                to_id = connection_match.group(3)
                to_label = connection_match.group(4) or to_id

                self.nodes[from_id] = from_label
                self.nodes[to_id] = to_label
                self.edges.append((from_id, to_id))
                continue

            # Standalone node definition: A[Label]
            node_match = re.search(r'(\w+)\[([^\]]+)\]', line)
            if node_match:
                node_id = node_match.group(1)
                label = node_match.group(2)
                self.nodes[node_id] = label

        # Validate that we found something
        if not self.nodes:
            raise ValueError("No valid Mermaid nodes found in diagram. Check syntax.")

        return self.nodes, self.edges


class FlowLayoutEngine:
    """Layout nodes based on flow hierarchy"""

    def __init__(self, nodes, edges, width, height):
        self.nodes = nodes
        self.edges = edges
        self.width = width
        self.height = height
        self.positions = {}

    def calculate_levels(self):
        """Determine hierarchical levels using BFS"""
        # Find root nodes (nodes with no incoming edges)
        incoming = defaultdict(int)
        outgoing = defaultdict(list)

        for from_id, to_id in self.edges:
            incoming[to_id] += 1
            outgoing[from_id].append(to_id)

        # Roots are nodes with no incoming edges
        roots = [node_id for node_id in self.nodes if incoming[node_id] == 0]

        # If no clear roots, use all nodes as potential starts
        if not roots:
            roots = list(self.nodes.keys())

        # BFS to assign levels
        levels = {}
        queue = deque([(root, 0) for root in roots])
        visited = set()

        while queue:
            node_id, level = queue.popleft()
            if node_id in visited:
                continue
            visited.add(node_id)
            levels[node_id] = level

            for child in outgoing[node_id]:
                if child not in visited:
                    queue.append((child, level + 1))

        # Assign level 0 to any unvisited nodes
        for node_id in self.nodes:
            if node_id not in levels:
                levels[node_id] = 0

        return levels

    def layout(self):
        """Calculate positions based on hierarchical levels"""
        levels = self.calculate_levels()

        # Group nodes by level
        level_groups = defaultdict(list)
        for node_id, level in levels.items():
            level_groups[level].append(node_id)

        max_level = max(levels.values()) if levels else 0

        # Calculate positions
        usable_width = self.width - 2 * MARGIN
        usable_height = self.height - 2 * MARGIN

        for level in range(max_level + 1):
            nodes_at_level = level_groups[level]
            num_nodes = len(nodes_at_level)

            if num_nodes == 0:
                continue

            # Vertical position based on level
            y_pos = MARGIN + (level / (max_level + 1)) * usable_height + SHAPE_HEIGHT / 2

            # Horizontal distribution
            if num_nodes == 1:
                x_positions = [self.width / 2]
            else:
                x_spacing = usable_width / (num_nodes + 1)
                x_positions = [MARGIN + (i + 1) * x_spacing for i in range(num_nodes)]

            for i, node_id in enumerate(nodes_at_level):
                self.positions[node_id] = (x_positions[i], y_pos)

        return self.positions


class HilbertLayoutEngine:
    """Layout nodes using Hilbert space-filling curve"""

    def __init__(self, nodes, edges, width, height):
        self.nodes = nodes
        self.edges = edges
        self.width = width
        self.height = height
        self.positions = {}

    def hilbert_d2xy(self, n, d):
        """Convert Hilbert curve distance to x,y coordinates"""
        x = y = 0
        s = 1
        while s < n:
            rx = 1 & (d // 2)
            ry = 1 & (d ^ rx)
            x, y = self.hilbert_rot(s, x, y, rx, ry)
            x += s * rx
            y += s * ry
            d //= 4
            s *= 2
        return x, y

    def hilbert_rot(self, n, x, y, rx, ry):
        """Rotate/flip quadrant appropriately"""
        if ry == 0:
            if rx == 1:
                x = n - 1 - x
                y = n - 1 - y
            x, y = y, x
        return x, y

    def layout(self):
        """Calculate positions using Hilbert curve"""
        num_nodes = len(self.nodes)

        # Find the smallest power of 2 that can contain all nodes
        n = 2 ** math.ceil(math.log2(math.sqrt(num_nodes)))

        usable_width = self.width - 2 * MARGIN - SHAPE_WIDTH
        usable_height = self.height - 2 * MARGIN - SHAPE_HEIGHT

        # Map each node to a Hilbert curve position
        for i, node_id in enumerate(self.nodes.keys()):
            hx, hy = self.hilbert_d2xy(n, i)

            # Scale to page dimensions
            x = MARGIN + SHAPE_WIDTH / 2 + (hx / n) * usable_width
            y = MARGIN + SHAPE_HEIGHT / 2 + (hy / n) * usable_height

            self.positions[node_id] = (x, y)

        return self.positions


class VisioGenerator:
    """Generate Visio document with shapes"""

    def __init__(self, layout_engine='flow',
                 horizontal_connections=DEFAULT_HORIZONTAL_CONNECTIONS,
                 vertical_connections=DEFAULT_VERTICAL_CONNECTIONS):
        self.layout_engine = layout_engine
        self.horizontal_connections = horizontal_connections
        self.vertical_connections = vertical_connections
        self.visio = None
        self.doc = None
        self.page = None
        self.master_shapes = {}  # Store master shapes by type

    def create_document(self):
        """Initialize Visio application and document"""
        self.visio = win32com.client.Dispatch("Visio.Application")
        self.visio.Visible = True

        # Create new document with Basic Diagram template
        self.doc = self.visio.Documents.Add("")
        self.page = self.doc.Pages.Item(1)

        # Set page size to A4
        self.page.PageSheet.CellsSRC(1, 0, 0).FormulaU = f"{A4_WIDTH} in"  # Width
        self.page.PageSheet.CellsSRC(1, 0, 1).FormulaU = f"{A4_HEIGHT} in"  # Height

    def add_connection_points(self, shape, horizontal_count, vertical_count):
        """
        Add connection points to a shape
        Based on your VBA code - adds points on all four sides
        """
        visSectionConnectionPts = 9
        visRowConnectionPts = 2
        visTagCnnctPt = 153
        visX = 0
        visY = 1

        # Ensure connection points section exists
        if not shape.SectionExists(visSectionConnectionPts, 1):
            shape.AddSection(visSectionConnectionPts)

        # Add vertical connection points (left and right sides)
        for i in range(1, vertical_count + 1):
            for side in [0, 1]:  # 0 = left (width*0), 1 = right (width*1)
                row_num = shape.AddRow(visSectionConnectionPts, visRowConnectionPts, visTagCnnctPt)

                # X position: left (0) or right (1)
                shape.CellsSRC(visSectionConnectionPts, row_num, visX).FormulaU = f"=width*{side}"

                # Y position: distributed vertically
                shape.CellsSRC(visSectionConnectionPts, row_num, visY).FormulaU = \
                    f"={i}*(height/{vertical_count}) - Height/{2 * vertical_count}"

        # Add horizontal connection points (top and bottom sides)
        for i in range(1, horizontal_count + 1):
            for side in [0, 1]:  # 0 = bottom (height*0), 1 = top (height*1)
                row_num = shape.AddRow(visSectionConnectionPts, visRowConnectionPts, visTagCnnctPt)

                # Y position: bottom (0) or top (1)
                shape.CellsSRC(visSectionConnectionPts, row_num, visY).FormulaU = f"=height*{side}"

                # X position: distributed horizontally
                shape.CellsSRC(visSectionConnectionPts, row_num, visX).FormulaU = \
                    f"={i}*(width/{horizontal_count}) - width/{2 * horizontal_count}"

    def create_master_shape(self, shape_type="Rectangle"):
        """
        Create a master shape with thicker border
        All other shapes of this type will reference this master for sizing
        """
        if shape_type in self.master_shapes:
            return self.master_shapes[shape_type]

        # Use Basic Shapes stencil
        basic_stencil = self.visio.Documents.OpenEx(
            self.visio.GetBuiltInStencilFile(3, 1),  # visBuiltInStencilBasicShapes
            4  # visOpenRO (read-only)
        )

        rectangle_master = basic_stencil.Masters(shape_type)

        # Place master shape in upper-left corner (off to the side)
        master_x = MARGIN
        master_y = A4_HEIGHT - MARGIN

        master_shape = self.page.Drop(rectangle_master, master_x, master_y)
        master_shape.Text = f"MASTER-{shape_type}"

        # Set shape size
        master_shape.CellsSRC(1, 1, 2).FormulaU = f"{SHAPE_WIDTH} in"  # Width
        master_shape.CellsSRC(1, 1, 3).FormulaU = f"{SHAPE_HEIGHT} in"  # Height

        # Make border thicker (LineWeight)
        master_shape.CellsU("LineWeight").FormulaU = "3 pt"

        # Add connection points to master
        self.add_connection_points(master_shape, self.horizontal_connections, self.vertical_connections)

        # Store the master shape
        self.master_shapes[shape_type] = master_shape

        print(f"Created master shape: {master_shape.Name} (ID: {master_shape.ID})")

        return master_shape

    def link_shape_to_master(self, shape, master_shape):
        """
        Link a shape's width and height to the master shape using GUARD formulas
        This replicates your SetMasterShape VBA function
        """
        master_name = master_shape.Name

        # Use GUARD to force the formulas and prevent manual override
        shape.Cells("Width").FormulaForce = f"GUARD({master_name}!Width)"
        shape.Cells("Height").FormulaForce = f"GUARD({master_name}!Height)"

    def create_shapes(self, nodes, positions):
        """Create Visio shapes at calculated positions"""
        # Create master shape first
        master_shape = self.create_master_shape("Rectangle")

        # Use Basic Shapes stencil - Rectangle
        basic_stencil = self.visio.Documents.OpenEx(
            self.visio.GetBuiltInStencilFile(3, 1),  # visBuiltInStencilBasicShapes
            4  # visOpenRO (read-only)
        )

        rectangle_master = basic_stencil.Masters("Rectangle")

        shape_map = {}
        for node_id, label in nodes.items():
            if node_id in positions:
                x, y = positions[node_id]

                # Drop shape on page
                shape = self.page.Drop(rectangle_master, x, y)
                shape.Text = label

                # Link to master shape for sizing
                self.link_shape_to_master(shape, master_shape)

                # Add connection points
                self.add_connection_points(shape, self.horizontal_connections, self.vertical_connections)

                shape_map[node_id] = shape

        return shape_map

    def generate(self, mermaid_text):
        """Main generation function"""
        # Parse Mermaid
        parser = MermaidParser(mermaid_text)
        nodes, edges = parser.parse()

        print(f"Parsed {len(nodes)} nodes and {len(edges)} edges")

        # Calculate layout
        if self.layout_engine == 'flow':
            layout = FlowLayoutEngine(nodes, edges, A4_WIDTH, A4_HEIGHT)
        else:
            layout = HilbertLayoutEngine(nodes, edges, A4_WIDTH, A4_HEIGHT)

        positions = layout.layout()

        # Create Visio document
        self.create_document()

        # Create shapes
        shapes = self.create_shapes(nodes, positions)

        print(f"Created {len(shapes)} shapes in Visio")
        print(f"Connection points: {self.horizontal_connections} horizontal, {self.vertical_connections} vertical")
        print("All shapes are linked to master shape for consistent sizing")
        print("Ready for manual connection in Visio!")

        return shapes


def load_from_file(filepath):
    """Load Mermaid diagram from file"""
    if not os.path.exists(filepath):
        print(f"ERROR: File '{filepath}' does not exist", file=sys.stderr)
        sys.exit(1)

    if not os.path.isfile(filepath):
        print(f"ERROR: '{filepath}' is not a file", file=sys.stderr)
        sys.exit(1)

    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()

        if not content.strip():
            print(f"ERROR: File '{filepath}' is empty", file=sys.stderr)
            sys.exit(1)

        print(f"Loaded Mermaid diagram from: {filepath}")
        return content

    except Exception as e:
        print(f"ERROR: Failed to read file '{filepath}': {e}", file=sys.stderr)
        sys.exit(1)


def load_from_clipboard():
    """Load Mermaid diagram from clipboard"""
    try:
        content = pyperclip.paste()

        if not content or not content.strip():
            print("ERROR: Clipboard is empty", file=sys.stderr)
            sys.exit(1)

        print("Loaded Mermaid diagram from clipboard")
        return content

    except Exception as e:
        print(f"ERROR: Failed to read from clipboard: {e}", file=sys.stderr)
        print("Make sure pyperclip is installed: pip install pyperclip", file=sys.stderr)
        sys.exit(1)


def main():
    """Main entry point with argument parsing"""
    parser = argparse.ArgumentParser(
        description='Convert Mermaid diagrams to Visio with automatic connection points',
        epilog="""
Examples:
  %(prog)s --file diagram.mmd
  %(prog)s --clipboard --layout hilbert
  %(prog)s --file diagram.txt --horizontal 7 --vertical 5
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    # Input source (mutually exclusive)
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument('--file', '-f',
                             help='Path to file containing Mermaid diagram')
    input_group.add_argument('--clipboard', '-c', action='store_true',
                             help='Read Mermaid diagram from clipboard')

    # Layout options
    parser.add_argument('--layout', '-l',
                        choices=['flow', 'hilbert'],
                        default='flow',
                        help='Layout algorithm (default: flow)')

    # Connection point options
    parser.add_argument('--horizontal', type=int,
                        default=DEFAULT_HORIZONTAL_CONNECTIONS,
                        help=f'Number of horizontal connection points (default: {DEFAULT_HORIZONTAL_CONNECTIONS})')

    parser.add_argument('--vertical', type=int,
                        default=DEFAULT_VERTICAL_CONNECTIONS,
                        help=f'Number of vertical connection points (default: {DEFAULT_VERTICAL_CONNECTIONS})')

    args = parser.parse_args()

    # Validate connection point counts
    if args.horizontal < 1 or args.horizontal > 20:
        print("ERROR: Horizontal connection points must be between 1 and 20", file=sys.stderr)
        sys.exit(1)

    if args.vertical < 1 or args.vertical > 20:
        print("ERROR: Vertical connection points must be between 1 and 20", file=sys.stderr)
        sys.exit(1)

    # Load Mermaid diagram
    if args.file:
        mermaid_text = load_from_file(args.file)
    else:
        mermaid_text = load_from_clipboard()

    # Display what we're doing
    print(f"\nConfiguration:")
    print(f"  Layout engine: {args.layout}")
    print(f"  Horizontal connection points: {args.horizontal}")
    print(f"  Vertical connection points: {args.vertical}")
    print()

    # Generate Visio diagram
    try:
        generator = VisioGenerator(
            layout_engine=args.layout,
            horizontal_connections=args.horizontal,
            vertical_connections=args.vertical
        )

        generator.generate(mermaid_text)
        print("\n✓ Successfully generated Visio diagram!")

    except ValueError as e:
        print(f"\nERROR: Invalid Mermaid syntax - {e}", file=sys.stderr)
        print("\nExpected format example:", file=sys.stderr)
        print("  graph TD", file=sys.stderr)
        print("    A[Node A] --> B[Node B]", file=sys.stderr)
        print("    B --> C[Node C]", file=sys.stderr)
        sys.exit(1)

    except Exception as e:
        print(f"\nERROR: Failed to generate diagram: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()