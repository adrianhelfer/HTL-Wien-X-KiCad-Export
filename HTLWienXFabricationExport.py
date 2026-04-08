import pcbnew
import os
import subprocess
import sys
import shutil

def find_kicad_cli():
    """
    Locates the kicad-cli executable in an OS-agnostic way.

    Returns:
        str: The absolute path to the kicad-cli executable, or None if not found.
    """
    # --- Check if kicad-cli is in the system's PATH ---
    kicad_cli_path = shutil.which("kicad-cli")
    if kicad_cli_path:
        return kicad_cli_path

    # --- If not in PATH, check default installation locations ---
    
    # Determine the executable name based on the OS
    cli_executable_name = "kicad-cli.exe" if sys.platform == "win32" else "kicad-cli"

    # Define platform-specific search paths
    search_paths = []
    if sys.platform == "darwin":  # macOS
        search_paths.append('/usr/local/bin/kicad-cli')
        search_paths.append(f"/Applications/KiCad.app/Contents/MacOS/{cli_executable_name}")
        # Check for versioned app names, e.g., KiCad-7.0.app
        for item in os.listdir("/Applications"):
             if item.lower().startswith("kicad") and item.endswith(".app"):
                  search_paths.append(f"/Applications/{item}/Contents/MacOS/{cli_executable_name}")

    elif sys.platform == "win32":  # Windows
        program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
        # Check for KiCad directory, which often contains a versioned subfolder
        kicad_base_dir = os.path.join(program_files, "KiCad")
        if os.path.isdir(kicad_base_dir):
            # Look inside version subdirectories (e.g., "8.0", "7.0")
            for version_dir in os.listdir(kicad_base_dir):
                potential_path = os.path.join(kicad_base_dir, version_dir, "bin", cli_executable_name)
                search_paths.append(potential_path)

    elif sys.platform.startswith("linux"): # Linux
        # Most package manager installs will be in the PATH.
        # This is a fallback for non-standard or manual installations.
        search_paths.append(f"/usr/bin/{cli_executable_name}")
        search_paths.append(f"/usr/local/bin/{cli_executable_name}")

    # Check the potential paths for an existing, executable file
    for path in search_paths:
        if os.path.isfile(path) and os.access(path, os.X_OK):
            return path
            
    return None

class HTLWienXFabricationExport(pcbnew.ActionPlugin):
    def defaults(self):
        self.name = "Fertigungsunterlagen exportieren"
        self.category = "Fabrication"
        self.description = "Exportiert die Layer F.Cu und B.Cu als SVG und die Bohrinformationen als EXC."
        self.show_toolbar_button = True
        self.icon_file_name = os.path.join(os.path.dirname(__file__), 'HTLWienX_24x24.png')

    def Run(self):
        kicad_cli_path = str(find_kicad_cli())
        board = pcbnew.GetBoard()
        board_path = board.GetFileName()
        board_directory = os.path.dirname(board_path)

        fabrication_directory = os.path.join(board_directory, 'Fabrication')
        if not os.path.exists(fabrication_directory):
            os.makedirs(fabrication_directory)

        output_filename = os.path.join(fabrication_directory, os.path.splitext(os.path.basename(board_path))[0])

        # Export B.Cu layer as SVG, black and white, the page as big as the board itself and not the drawing sheet (page size mode 2), exclude drawing sheet, negative, named {filename}-Bottom.svg, in a directory called /Fabrication/ located in the same directory as the board file
        # F.Cu is mirrored so the ink is closer to the copper
        fCuExportCommand = [
            kicad_cli_path,
            "pcb", "export", "svg",
            "-o", f"{output_filename}-Top.svg",
            "--layers", "F.Cu",
            "--mirror",
            "-n",
            "--black-and-white",
            "--page-size-mode", "2",
            "--exclude-drawing-sheet",
            board_path
        ]

        bCuExportCommand = [
            kicad_cli_path,
            "pcb", "export", "svg",
            "-o", f"{output_filename}-Bottom.svg",
            "--layers", "B.Cu",
            "-n",
            "--black-and-white",
            "--page-size-mode", "2",
            "--exclude-drawing-sheet",
            board_path
        ]

        drillExportCommand = [
            kicad_cli_path,
            "pcb", "export", "drill",
            "-o", fabrication_directory,
            "--drill-origin", "plot",
            "--excellon-zeros-format", "suppressleading",
            "-u", "in",
            "--excellon-min-header",
            board_path
        ]

        subprocess.run(bCuExportCommand, check=True)
        subprocess.run(fCuExportCommand, check=True)
        subprocess.run(drillExportCommand, check=True)

        # Split DRL to EXC and Tool Info File
        drl_file = os.path.join(fabrication_directory, os.path.splitext(os.path.basename(board_path))[0] + '.drl')
        exc_file = os.path.join(fabrication_directory, os.path.splitext(os.path.basename(board_path))[0] + '.exc')
        tool_file = os.path.join(fabrication_directory, os.path.splitext(os.path.basename(board_path))[0] + '-Bohrer.txt')

        nc_file = os.path.join(fabrication_directory, os.path.splitext(os.path.basename(board_path))[0] + '.NC')
        # Make nc_file content ansi compatible
        nc_file = nc_file.encode('ascii', 'ignore').decode('ascii')

        tool_diameters = []
        with open(drl_file, 'r') as drl:
            # Skip to tool list
            while True:
                line = drl.readline()
                if line.startswith('T1'):
                    break
            
            # Read tool list
            while '%' not in line:
                parts = line.strip().split('C')
                tool_id = parts[0]  # e.g., T1
                diameter_inch = float(parts[1])  # e.g., 0.0315
                diameter_mm = diameter_inch * 25.4  # Convert to mm
                tool_diameters.append((tool_id, diameter_mm))
                line = drl.readline()

            # Sort by diameter (2nd element of tuple)
            tool_diameters.sort(key=lambda x: x[1])

            # Write coordinates to EXC file with absolute coordinates
            with open(exc_file, 'w') as exc:
                while True:
                    line = drl.readline()
                    if not line:
                        break
                    if 'G' not in line:
                        exc.write(line.replace('Y-', 'Y').replace('X-', 'X'))
        
        # Write tool info to tool file
        with open(tool_file, 'w') as tool:
            tool.write('Bohrplotter        OG-ID\n========================\n')
            i = 1
            missing_tool_warning = False
            for tool_id, diameter_mm in tool_diameters:
                writeline = f'T{i:03.0f} {diameter_mm:4.1f} mm '
                if diameter_mm < 0.8:
                    writeline += '(!)'
                    missing_tool_warning = True
                else:
                    writeline += '   '
                
                writeline += f'   {tool_id}\n'

                tool.write(writeline)

                i += 1
            
            if missing_tool_warning:
                tool.write('\n(!) Nicht im Sortiment\n')
        
        # NC Processor

        nc_content_without_line_numbers = []

        def drehzahl(durchmesser):
            if durchmesser == 0:
                return 0
            ret = 12609.69 / durchmesser**1.0236
            # chatGPTs Antwort auf die Wertetabelle von PrimCAM
            return int(round(ret, -2))

        def vorschub(durchmesser):
            return int(drehzahl(durchmesser) / 20)

        def drehzahl_vorschub(durchmesser):
            return (round(drehzahl(durchmesser), -2), round(vorschub(durchmesser), -2))

        def neues_werkzeug(werkzeugnummer, durchmesser):
            durchmesser = round(durchmesser, 2)
            nc_content_without_line_numbers.append(f"(BOHREN ø{durchmesser})")
            nc_content_without_line_numbers.append(f"{werkzeugnummer} M09 (Bohrer ø{durchmesser})")
            nc_content_without_line_numbers.append("M06")

        def nc_koordinate(exc_koordinate):
            ret = (exc_koordinate * 25.4) / 10000
            return f'{ret:0.3f}'

        def spin_up(nc_x, nc_y, durchmesser):
            nc_content_without_line_numbers.append(f'G00 X{nc_x} Y{nc_y} Z5')
            nc_content_without_line_numbers.append(f'S{drehzahl(durchmesser)} M13')
            nc_content_without_line_numbers.append(f'G81 X{nc_x} Y{nc_y} Z-4 R5 F{vorschub(durchmesser)}')

        tool_diameters = []
        exc_drill_infos = {}
        nc_drill_infos = {}
        with open(drl_file, 'r') as drl:
            # Skip to tool list
            while True:
                line = drl.readline()
                if line.startswith('T1'):
                    break
            
            # Read tool list
            while '%' not in line:
                parts = line.strip().split('C')
                tool_id = parts[0]  # e.g., T1
                diameter_inch = float(parts[1])  # e.g., 0.0315
                diameter_mm = diameter_inch * 25.4  # Convert to mm
                tool_diameters.append((tool_id, diameter_mm))
                line = drl.readline()

            # Sort by diameter (2nd element of tuple)
            tool_diameters.sort(key=lambda x: x[1])

            # Skip to drill coordinates list for specific tool
            while True:
                line = drl.readline()
                if line.startswith('T'):
                    break
            
            while line:
                if 'G' in line or 'M' in line:
                    line = drl.readline()
                    continue
                tool_index = line[:-1]  # e.g., T1
                line = drl.readline()
                if 'G' in line or 'M' in line:
                    line = drl.readline()
                    continue
                exc_drill_infos[tool_index] = []
                while 'T' not in line and 'X' in line and 'Y' in line and 'M' not in line:
                    exc_koordinates = line.split('Y')
                    exc_x = exc_koordinates[0][1:]
                    exc_y = exc_koordinates[1][:-1]
                    exc_drill_infos[tool_index].append((exc_x, exc_y))
                    line = drl.readline()

        # Iterate over tool_diameters and match with exc_drill_infos
        for tool_id, diameter in tool_diameters:
            nc_drill_infos[tool_id] = []
            if tool_id not in exc_drill_infos:
                continue
            for exc_koordinate in exc_drill_infos[tool_id]:
                nc_x = nc_koordinate(float(exc_koordinate[0]))
                nc_y = nc_koordinate(float(exc_koordinate[1]))
                nc_drill_infos[tool_id].append((nc_x, nc_y))


        nc_program_name = ''
        # Give nc_program_name the name of the drl file without the extension and path
        nc_program_name = drl_file.split('\\')[-1].split('.')[0]
        # Make nc_program_name ansi compatible
        nc_program_name = nc_program_name.encode('ascii', 'ignore').decode('ascii')


        nc_content_without_line_numbers.append(f'({nc_program_name})')
        nc_content_without_line_numbers.append('G54 G90 G17')


        # Iterate over tool_diameters and match with nc_drill_infos and print only the coordinates that have changed from the previous line
        for tool_id, diameter in tool_diameters:
            if tool_id not in exc_drill_infos:
                continue
            neues_werkzeug(tool_id, diameter)
            prev_x = -1
            prev_y = -1
            first_inner_loop = True
            for nc_koordinate in nc_drill_infos[tool_id]:
                if first_inner_loop == True:
                    spin_up(nc_koordinate[0], nc_koordinate[1], diameter)
                    first_inner_loop = False
                    continue
                if nc_koordinate[0] != prev_x:
                    nc_content_without_line_numbers.append(f'X{nc_koordinate[0]}')
                    prev_x = nc_koordinate[0]
                if nc_koordinate[1] != prev_y:
                    nc_content_without_line_numbers.append(f'Y{nc_koordinate[1]}')
                    prev_y = nc_koordinate[1]
            nc_content_without_line_numbers.append('G80')
        nc_content_without_line_numbers.append('M30')

        with open(nc_file, 'w', encoding="latin_1") as nc:
            line_counter = 1
            for line in nc_content_without_line_numbers:
                nc.write(f'N{line_counter:04.0f} {line}\n'.replace('Y-', 'Y').replace('X-', 'X'))
                line_counter += 1

        os.remove(drl_file)

HTLWienXFabricationExport().register() # Instantiate and register to Pcbnew
