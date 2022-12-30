from pathlib import Path

import click
import pandas as pd
from zahner_analysis.analysis_tools.eis_fitting import EisFitting, EisFittingResult
from zahner_analysis.file_import.impedance_model_import import IsfxModelImport
from zahner_analysis.file_import.ism_import import IsmImport


def fit_model(
    model_path: Path, z_data_path: Path, fitting_conn: EisFitting = EisFitting()
) -> tuple[str, EisFittingResult]:
    """Fits impedance data to given circuit model and returns result

    Arguments:
        model_path (pathlib.Path): path to *.isfx model to use for fitting
        z_data_path (pathlib.Path): path to *.ism impedance dataset
        fitting_conn (EisFitting): optional fitting object which is used to interact with
            the Zahner Analysis program using its REST API. If a fitting object is not
            given, a new one will be automatically created.

    Returns:
        str: filename, excluding extension, of raw impedance data file
        EisFittingResult: results of impedance model fitting
    """
    if model_path.exists():
        model = IsfxModelImport(model_path)
    else:
        raise FileNotFoundError("Model path does not exist")

    if model_path.exists():
        z_data = IsmImport(z_data_path)
    else:
        raise FileNotFoundError("Impedance data path does not exist")

    result = fitting_conn.fit(model, z_data)
    return z_data_path.stem, result


def fit_directory_of_models(
    model_path: Path, z_data_dir: Path
) -> dict[str, EisFittingResult]:
    """Will fit entire directory of impedance datasets to the given model

    Arguments:
        model_path (pathlib.Path): path to *.isfx model to use for fitting
        z_data_dir (pathlib.Path): path to directory containing one or more *.ism files

    Returns:
        dict[str, EisFittingResult]: dictionary containing fit results where the keys are
            the filenames of the *.ism files (excluding the extension) and the values are
            the fitting results
    """
    if not z_data_dir.exists():
        raise FileNotFoundError("Directory path does not exist")
    if not z_data_dir.is_dir():
        raise NotADirectoryError("Path is not a directory")

    fitting_conn = EisFitting()
    results = [
        fit_model(model_path, z_data_path, fitting_conn)
        for z_data_path in z_data_dir.glob("*.ism")
    ]
    return dict(results)


def gen_result_table_row(
    filename: str, result_output: EisFittingResult
) -> pd.DataFrame:
    """Generates a dataframe containing fitting results for a single result

    Arguments:
        filename (str): filename to use as value in first column of dataframe to identify
            result
        result_output (EisFittingResult): fitting results from which to extract data into
            data in the dataframe

    Returns:
        pd.DataFrame: dataframe containing flattened (normalized) fitting result data
            including fitted model and overall metadata about the fit. Can be
            concatenated with other similar dataframes to form complete bulk dataset
    """
    results: pd.DataFrame = pd.json_normalize(result_output.getFitResultJson())
    results.insert(0, "filename", filename)
    results = results.apply(pd.to_numeric, errors="ignore")
    return results


@click.command()
@click.option(
    "--model_path",
    "-m",
    help="Path to model of type *.isfx",
    type=click.Path(path_type=Path),
)
@click.option(
    "--z_data_dir",
    "-z",
    help="Path to directory containing impedance data in form of one or more *.ism files",
    type=click.Path(path_type=Path),
)
@click.option(
    "--output_filename",
    "-o",
    help='Filename to use for output file. Default is "output".',
    type=click.STRING,
    default="output",
)
def cli(model_path: Path, z_data_dir: Path, output_filename: str):
    # This docstring is used to display --help output for the cli, so does not follow
    # typical python docstring format
    """
    Small command line tool to bulk fit Zahner electrochemical impedance spectroscopy
    datasets in a single directory to a single circuit model (*.isfx file). Results of all the fits
    will be compiled and written to a Microsoft Excel document (*.xlsx file) in the same
    directory as the raw impedance datasets (*.ism files).
    """
    results = fit_directory_of_models(model_path, z_data_dir)
    table = pd.concat(
        [gen_result_table_row(filename, result) for filename, result in results.items()]
    )
    table.to_excel(z_data_dir / f"{output_filename}.xlsx", index=False)


if __name__ == "__main__":
    cli.main()
