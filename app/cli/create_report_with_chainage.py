from app.utils.functions.util_functions import concurrent_ASFT
from app.utils.report import write_report_with_chainage
from pathlib import Path as pathlib_Path
from inquirer import prompt, List, Path, Text
import click
from yaspin import yaspin


def run_write_report(left_file, right_file, weather, runway_material, runway_length, starting_point, output_folder):
    L, R = concurrent_ASFT(left_file, right_file)
    L.weather = weather
    L.runway_material = runway_material
    L.runway_length = int(runway_length)
    L.starting_point = int(starting_point)

    write_report_with_chainage(L, R, L.runway_length, L.starting_point, output_folder)


@click.command()
def main():
    left_file_question = Path("left_file", message="Medición lado izquierdo:", exists=True)
    right_file_question = Path("right_file", message="Medición lado derecho:", exists=True)
    output_folder_question = Path("output_folder", message="Carpeta de destino", exists=True)
    weather_question = List(
        "weather",
        message="Condición metereológica:",
        choices=["Bueno", "Nublado", "Soleado", "Lluvioso", "Escarcha"],
    )
    runway_material_question = List("runway_material", message="Tipo de pavimento:", choices=["Asfalto", "Hormigón"])
    runway_length_question = Text("runway_length", message="Longitud de la pista (múltiplos de 10):")
    starting_point_question = Text("starting_point", message="Punto de inicio (múltiplos de 10):")

    answers = prompt(
        [
            left_file_question,
            right_file_question,
            output_folder_question,
            weather_question,
            runway_material_question,
            runway_length_question,
            starting_point_question,
        ]
    )

    left_file = pathlib_Path(answers["left_file"]).resolve()
    right_file = pathlib_Path(answers["right_file"]).resolve()
    weather = answers["weather"]
    runway_material = answers["runway_material"]
    output_folder = pathlib_Path(answers["output_folder"]).resolve()
    runway_length = answers["runway_length"]
    starting_point = answers["starting_point"]

    with yaspin(text="Cargando...", spinner="line") as spinner:
        run_write_report(left_file, right_file, weather, runway_material, runway_length, starting_point, output_folder)
        spinner.text = "¡Listo!"
        spinner.ok("✓")


if __name__ == "__main__":
    main()
