
import questionary
from .cli import app

def main():
    """Main function for the interactive menu."""
    while True:
        choice = questionary.select(
            "What do you want to do?",
            choices=[
                "Compile a report",
                "Convert PDF to SVG",
                "Manage Word integration",
                "Exit"
            ]
        ).ask()

        if choice == "Compile a report":
            handle_compile()
        elif choice == "Convert PDF to SVG":
            handle_svg_import()
        elif choice == "Manage Word integration":
            handle_word_integration()
        elif choice == "Exit":
            break

def handle_compile():
    """Handle the compile command."""
    input_file = questionary.path("Enter the path to the input DOCX file:").ask()
    output_file = questionary.path("Enter the path to the output PDF file:").ask()
    keep_temp = questionary.confirm("Keep temporary files?", default=False).ask()
    verbose = questionary.confirm("Enable verbose logging?", default=False).ask()
    log_file = questionary.path("Enter the path to the log file (optional):", default="").ask()

    args = ["compile", input_file, output_file]
    if keep_temp:
        args.append("--keep-temp")
    if verbose:
        args.append("--verbose")
    if log_file:
        args.extend(["--log-file", log_file])

    try:
        app(args, standalone_mode=False)
    except SystemExit:
        pass

def handle_svg_import():
    """Handle the svg-import command."""
    input_file = questionary.path("Enter the path to the input PDF file:").ask()
    output_file = questionary.path("Enter the path to the output SVG file:").ask()
    page = questionary.text("Enter the page(s) to convert (e.g., 1, 1-3, all):", default="all").ask()
    verbose = questionary.confirm("Enable verbose logging?", default=False).ask()
    log_file = questionary.path("Enter the path to the log file (optional):", default="").ask()

    args = ["svg-import", input_file, output_file, "--page", page]
    if verbose:
        args.append("--verbose")
    if log_file:
        args.extend(["--log-file", log_file])

    try:
        app(args, standalone_mode=False)
    except SystemExit:
        pass

def handle_word_integration():
    """Handle the word-integration command."""
    choice = questionary.select(
        "What do you want to do?",
        choices=[
            "Install",
            "Remove",
            "Update",
            "Status",
            "Back"
        ]
    ).ask()

    if choice == "Back":
        return

    args = ["word-integration", choice.lower()]

    try:
        app(args, standalone_mode=False)
    except SystemExit:
        pass
