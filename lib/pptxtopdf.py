import os
import comtypes
import comtypes.client


def pptx_to_pdf(input_path, output_path):

    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    comtypes.CoInitialize()

    powerpoint = None
    presentation = None

    try:
        powerpoint = comtypes.client.CreateObject(
            "PowerPoint.Application"
        )

        powerpoint.Visible = True

        presentation = powerpoint.Presentations.Open(
            input_path,
            ReadOnly=True,
            Untitled=True,
            WithWindow=False
        )

        # 32 = PDF
        presentation.SaveAs(
            output_path,
            32
        )

    finally:

        if presentation is not None:
            presentation.Close()

        if powerpoint is not None:
            powerpoint.Quit()

        comtypes.CoUninitialize()