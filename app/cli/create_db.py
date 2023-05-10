from app.utils.functions.util_functions import concurrent_ASFT
from app.utils.excel_db import add_data_to_db

from pathlib import Path as pathlib_Path
from inquirer import prompt, List, Path, Text
import click
from yaspin import yaspin


@click.command()
def main():
    folder_path_question = Path("folder_path", message="Carpeta de mediciones", exists=True)
    db_file_question = Path("db_file", message="Archivo de base de datos")

    runway_length_question = Text("runway_length", message="Longitud de la pista (múltiplos de 10):")
    starting_point_1_question = Text(
        "starting_point_1", message="Punto de inicio para cabecera ∈ [01 - 18] (múltiplos de 10):"
    )
    starting_point_2_question = Text(
        "starting_point_2", message="Punto de inicio para cabecera ∈ [19 - 36] (múltiplos de 10):"
    )
    operator_question = Text("operator", message="Operador:")
    temperature_question = Text("temperature", message="Temperatura:")
    surface_condition_question = List(
        "surface_condition", message="Condición de superficie:", choices=["Seco", "Húmedo"]
    )
    weather_question = List(
        "weather",
        message="Condición metereológica:",
        choices=["Bueno", "Nublado", "Soleado", "Lluvioso", "Escarcha"],
    )
    runway_material_question = List("runway_material", message="Tipo de pavimento:", choices=["Asfalto", "Hormigón"])

    answers = prompt(
        [
            folder_path_question,
            db_file_question,
            runway_length_question,
            starting_point_1_question,
            starting_point_2_question,
            operator_question,
            temperature_question,
            surface_condition_question,
            weather_question,
            runway_material_question,
        ]
    )

    folder_path = pathlib_Path(answers["folder_path"]).resolve()
    db_file = pathlib_Path(answers["db_file"]).resolve()
    measurements = folder_path

    file_list = []
    for item in measurements.iterdir():
        if item.is_file():
            file_list.append(item)

    with yaspin(text="Cargando...", spinner="line") as spinner:
        data = concurrent_ASFT(file_list)
        for i in data:
            i.operator = answers["operator"]
            i.temperature = int(answers["temperature"])
            i.surface_condition = answers["surface_condition"]
            i.weather = answers["weather"]
            i.runway_material = answers["runway_material"]
            i.runway_length = int(answers["runway_length"])
            try:
                if int(i.numbering) <= 18:
                    i.starting_point = int(answers["starting_point_1"])
                    add_data_to_db(i, db_file)
                else:
                    i.starting_point = int(answers["starting_point_2"])
                    add_data_to_db(i, db_file)
                print(f"Added {i.filename} to the database.")
            except Exception as e:
                print(f"Error processing {i.filename}: {e}")
                continue

        spinner.text = "¡Listo!"
        spinner.ok("✓")


if __name__ == "__main__":
    main()
