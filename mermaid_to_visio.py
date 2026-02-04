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

            # Skip diagram type declarations, empty lines, and comments
            if not line or line.startswith('graph') or line.startswith('flowchart') or line.startswith('%'):
                continue

            # Handle curly braces in node definitions (e.g., C{Authenticated?})
            # Replace with square brackets for consistency
            line = re.sub(r'(\w+)\{([^}]+)\}', r'\1[\2]', line)

            # Handle different arrow types: -->, --->, -.->
            # Normalize to -->
            line = re.sub(r'[-\.]{2,}>', '-->', line)

            # Parse connection with labels on edges: A -->|label| B or A --|label|--> B
            edge_label_match = re.search(
                r'(\w+)(?:\[([^\]]+)\])?\s*-->\|([^\|]+)\|\s*(\w+)(?:\[([^\]]+)\])?',
                line
            )
            if edge_label_match:
                from_id = edge_label_match.group(1)
                from_label = edge_label_match.group(2)
                # edge_label = edge_label_match.group(3)  # Edge label (ignored)
                to_id = edge_label_match.group(4)
                to_label = edge_label_match.group(5)

                # Only set label if we have one and node doesn't exist yet
                # CRITICAL: Don't overwrite existing labels!
                if from_label:
                    self.nodes[from_id] = from_label
                elif from_id not in self.nodes:
                    self.nodes[from_id] = from_id

                if to_label:
                    self.nodes[to_id] = to_label
                elif to_id not in self.nodes:
                    self.nodes[to_id] = to_id

                self.edges.append((from_id, to_id))
                continue

            # Parse standard connection: A --> B or A[Label] --> B[Label]
            connection_match = re.search(
                r'(\w+)(?:\[([^\]]+)\])?\s*-->\s*(\w+)(?:\[([^\]]+)\])?',
                line
            )
            if connection_match:
                from_id = connection_match.group(1)
                from_label = connection_match.group(2)
                to_id = connection_match.group(3)
                to_label = connection_match.group(4)

                # Only set label if we have one and node doesn't exist yet
                # CRITICAL: Don't overwrite existing labels!
                if from_label:
                    self.nodes[from_id] = from_label
                elif from_id not in self.nodes:
                    self.nodes[from_id] = from_id

                if to_label:
                    self.nodes[to_id] = to_label
                elif to_id not in self.nodes:
                    self.nodes[to_id] = to_id

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

        # Debug output
        print(f"Parsed nodes:")
        for node_id, label in self.nodes.items():
            print(f"  {node_id}: '{label}'")
        print(f"Parsed edges: {self.edges}")

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

            # Vertical position based on level (top to bottom)
            if max_level == 0:
                y_pos = self.height / 2
            else:
                y_pos = A4_HEIGHT - MARGIN - (level / max_level) * usable_height

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
        self.stencil = None
        self.rectangle_master = None

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

    def find_rectangle_master(self):
        """
        Find a suitable rectangle master from Basic Shapes stencil
        """
        if self.rectangle_master:
            return self.rectangle_master

        # Get the path to Basic Shapes stencil
        # Try multiple methods to find it
        stencil_paths = [
            "BASFLO_M.VSSX",  # Basic Flowchart Shapes (includes Rectangle/Process)
            "BASIC_U.VSSX",   # Basic Shapes
            "BASFLO_U.VSSX",  # Basic Flowchart
        ]

        for stencil_name in stencil_paths:
            try:
                # Try to open the stencil
                stencil_path = os.path.join(
                    os.path.dirname(self.visio.Path),
                    stencil_name
                )

                # Try direct stencil name first
                try:
                    basic_stencil = self.visio.Documents.OpenEx(
                        stencil_name,
                        4  # visOpenRO (read-only)
                    )
                    self.stencil = basic_stencil
                    print(f"Opened stencil: {stencil_name}")
                    break
                except:
                    # Try with full path
                    if os.path.exists(stencil_path):
                        basic_stencil = self.visio.Documents.OpenEx(
                            stencil_path,
                            4  # visOpenRO (read-only)
                        )
                        self.stencil = basic_stencil
                        print(f"Opened stencil: {stencil_path}")
                        break
            except Exception as e:
                continue

        if not self.stencil:
            print("⚠ Could not open any stencil, will use DrawRectangle")
            return None

        # Try various possible names for rectangle shape
        rectangle_names = [
            "Rectangle",
            "Process",  # From flowchart stencil
            "Box",
            "Square",
        ]

        for name in rectangle_names:
            try:
                master = self.stencil.Masters(name)
                print(f"✓ Found rectangle master: {name}")
                self.rectangle_master = master
                return master
            except:
                continue

        # If none found, use DrawRectangle as fallback
        print("⚠ No rectangle master found, will use DrawRectangle")
        return None

    def create_rectangle_shape(self, x, y):
        """
        Create a rectangle shape at given position
        """
        if self.rectangle_master is None:
            self.find_rectangle_master()

        if self.rectangle_master:
            # Use master to drop shape
            shape = self.page.Drop(self.rectangle_master, x, y)
        else:
            # Draw rectangle directly - coordinates are center point
            x1 = x - SHAPE_WIDTH / 2
            y1 = y - SHAPE_HEIGHT / 2
            x2 = x + SHAPE_WIDTH / 2
            y2 = y + SHAPE_HEIGHT / 2
            shape = self.page.DrawRectangle(x1, y1, x2, y2)

        return shape

    def add_connection_points(self, shape, horizontal_count, vertical_count):
        """
        Add connection points to a shape
        Based on your VBA code - adds points on all four sides
        """
        # Section and row constants
        visSectionConnectionPts = 9
        visRowConnectionPts = 0
        visTagCnnctPt = 153

        # Cell indices
        visX = 0
        visY = 1

        # CRITICAL: Check if section exists LOCALLY (parameter 1, not 0)
        # This is what your VBA does!
        if not shape.SectionExists(visSectionConnectionPts, 1):
            shape.AddSection(visSectionConnectionPts)

        # Add vertical connection points (left and right sides)
        for i in range(1, vertical_count + 1):
            for side in [0, 1]:  # 0 = left, 1 = right
                row_num = shape.AddRow(visSectionConnectionPts, visRowConnectionPts, visTagCnnctPt)

                # Use the returned row_num to set formulas
                shape.CellsSRC(visSectionConnectionPts, row_num, visX).Formula = f"=Width*{side}"
                shape.CellsSRC(visSectionConnectionPts, row_num, visY).Formula = \
                    f"={i}*(Height/{vertical_count})-Height/{2 * vertical_count}"

        # Add horizontal connection points (top and bottom sides)
        for i in range(1, horizontal_count + 1):
            for side in [0, 1]:  # 0 = bottom, 1 = top
                row_num = shape.AddRow(visSectionConnectionPts, visRowConnectionPts, visTagCnnctPt)

                shape.CellsSRC(visSectionConnectionPts, row_num, visY).Formula = f"=Height*{side}"
                shape.CellsSRC(visSectionConnectionPts, row_num, visX).Formula = \
                    f"={i}*(Width/{horizontal_count})-Width/{2 * horizontal_count}"

    def create_master_shape(self):
        """
        Create a master shape with thicker border
        All other shapes of this type will reference this master for sizing
        """
        if "master" in self.master_shapes:
            return self.master_shapes["master"]

        # Place master shape in upper-left corner
        master_x = MARGIN + SHAPE_WIDTH / 2
        master_y = A4_HEIGHT - MARGIN - SHAPE_HEIGHT / 2

        master_shape = self.create_rectangle_shape(master_x, master_y)
        master_shape.Text = "MASTER"

        # Use FormulaForce to override any guards on the cells
        try:
            master_shape.Cells("Width").FormulaForce = f"{SHAPE_WIDTH} in"
            master_shape.Cells("Height").FormulaForce = f"{SHAPE_HEIGHT} in"
        except Exception as e:
            print(f"⚠ Warning: Could not set master shape size: {e}")
            print("  Continuing with default size...")

        # Make border thicker (LineWeight)
        try:
            master_shape.Cells("LineWeight").FormulaU = "3 pt"
        except:
            pass

        # Add connection points to master
        self.add_connection_points(master_shape, self.horizontal_connections, self.vertical_connections)

        # Store the master shape
        self.master_shapes["master"] = master_shape

        print(f"✓ Created master shape: {master_shape.Name} (ID: {master_shape.ID})")

        return master_shape

    def link_shape_to_master(self, shape, master_shape):
        """
        Link a shape's width and height to the master shape using GUARD formulas
        This replicates your SetMasterShape VBA function
        """
        master_name = master_shape.Name

        try:
            # Use GUARD to force the formulas and prevent manual override
            shape.Cells("Width").FormulaForce = f"GUARD({master_name}!Width)"
            shape.Cells("Height").FormulaForce = f"GUARD({master_name}!Height)"
        except Exception as e:
            print(f"⚠ Warning: Could not link shape {shape.Text[:20]} to master: {e}")

    def create_shapes(self, nodes, positions):
        """Create Visio shapes at calculated positions"""
        # Create master shape first
        master_shape = self.create_master_shape()

        shape_map = {}
        for node_id, label in nodes.items():
            if node_id not in positions:
                continue

            x, y = positions[node_id]

            # Create shape
            shape = self.create_rectangle_shape(x, y)

            # Set text - this is critical!
            shape.Text = label
            print(f"Created shape '{node_id}' with label '{label}' at ({x:.2f}, {y:.2f})")

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

        print(f"\nCalculated positions:")
        for node_id, pos in positions.items():
            print(f"  {node_id}: {pos}")

        # Create Visio document
        self.create_document()

        # Create shapes
        print(f"\nCreating shapes...")
        shapes = self.create_shapes(nodes, positions)

        print(f"\n✓ Created {len(shapes)} shapes in Visio")
        print(f"✓ Connection points: {self.horizontal_connections} horizontal, {self.vertical_connections} vertical")
        print("✓ All shapes are linked to master shape for consistent sizing")
        print("\n🎯 Ready for manual connection in Visio!")

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
        print("\n✅ Successfully generated Visio diagram!")

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