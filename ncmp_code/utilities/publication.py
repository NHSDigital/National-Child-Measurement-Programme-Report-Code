import win32com.client as win32
import logging
import xlwings as xw


def save_charts_as_image(source_file, output_path, reportyear, image_type="png"):
    """
    Save updated charts as images to the final charts folder. The function
    will loop through all tabs / charts in the specified file and save each
    to the output folder.

    Parameters
    ----------
    source_file : path
        File location of the Excel file that contains the charts to be saved.
    image_type: str
        Type of image file to create. Tested for jpeg (jpg) and png file types.
        Default is png.
    output_path : path
        Folder location where the images should be saved.
    reportyear : str
        Report year images relate to - used in filename
    Returns
    -------
        None
    """
    logging.info("Saving final publication charts")

    # Activate the chart master file using win32 (to allow selection of chart objects)
    app = win32.gencache.EnsureDispatch("Excel.Application")
    app.WindowState = win32.constants.xlMaximized
    wb = app.Workbooks.Open(Filename=source_file)

    # Create a list of all the worksheets in the file
    sheet_names = [sheet.Name for sheet in wb.Sheets]

    # For each worksheet, write the chart as a the specified type, named as
    # per the worksheet.
    for sheet in sheet_names:
        sht = wb.Worksheets[sheet]
        for chartObject in sht.ChartObjects():
            output_file = sheet + "_" + reportyear + "." + image_type
            chartObject.Chart.Export(output_path / output_file)


def save_masterfile(source_file, output_path, save_name):
    """
    Save updated master file to the specified folder and rename

    Parameters
    ----------
    source_file : path
        filepath of the master Excel file that needs to be saved

    output_path: path
        filepath of where to save the final file

    save_name: str
        new name for final file

    Returns
    -------
        None
    """
    logging.info(
        f"Saving final version of master file as {save_name}"
        )

    # Select the master file
    xw.App()
    wb = xw.books.open(source_file)

    # Select first sheet of workbook
    sht = wb.sheets[0]
    sht.select()

    # Save the master file to the specified folder, named as per save_name input
    savepath = output_path / save_name
    wb.save(savepath)

    # Close Excel
    xw.apps.active.api.Quit()
