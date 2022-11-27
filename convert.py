import pathlib

import win32com
from win32com import client


def convert_to_pdf_libre(src: str, dst: str):
    def create_prop(name, value):
        prop = oo_service_manager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        prop.Name = name
        prop.Value = value
        return prop

    src = "file:///%s" % src
    dst = "file:///%s" % dst

    oo_service_manager = win32com.client.DispatchEx("com.sun.star.ServiceManager")
    desktop = oo_service_manager.CreateInstance("com.sun.star.frame.Desktop")
    oo_service_manager._FlagAsMethod("Bridge_GetStruct")

    loading_properties = [create_prop("ReadOnly", True), create_prop("Hidden", True)]

    document = desktop.loadComponentFromUrl(src, "_blank", 0, tuple(loading_properties))
    document.CurrentController.Frame.ContainerWindow.Visible = False

    pdf_properties = [create_prop("FilterName", "writer_pdf_Export")]

    document.storeToURL(dst, pdf_properties)
    document.close(True)


def convert_to_pdf_wps(src: str, dst: str):
    word = win32com.client.DispatchEx("Kwps.Application")
    word.Visible = False
    doc = word.Documents.Open(src)
    doc.ExportAsFixedFormat(dst, 17)
    doc.Close()
    word.Quit()


def convert_to_pdf_ms(src: str, dst: str):
    word = client.DispatchEx("Word.Application")
    doc = word.Documents.Open(src)
    doc.SaveAs(dst, 17)
    doc.Close()
    word.Quit()


def convert_to_pdf(src: str, dst: str):
    src = pathlib.PurePath(src).as_posix()
    dst = pathlib.PurePath(dst).as_posix()
    try:
        convert_to_pdf_ms(src, dst)
    except Exception:
        print("no ms office")
    else:
        print("use ms office convert to pdf")
        return
    try:
        convert_to_pdf_libre(src, dst)
    except Exception:
        print("no libre office")
    else:
        print("use libre office convert to pdf")
        return
    try:
        convert_to_pdf_wps(src, dst)
    except Exception:
        print("no wps office")
    else:
        print("use wps office convert to pdf")
        return

