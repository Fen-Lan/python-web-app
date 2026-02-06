import uno
from com.sun.star.beans import PropertyValue
import time
import os, traceback
# todo 要先把libreoffice的服务开启： env OOO_MAX_THREADS=8 soffice --headless --accept="socket,host=localhost,port=8100;urp;" -env:SOFFICE_MAX_MEM=4096 -env:SOFFICE_CACHE_SIZE=1024
# todo 也可以使用命令：soffice --headless --accept="socket,host=localhost,port=8100;urp;" --nofirststartwizard

class LibreExcel:
    def __init__(self):
        """
        Initialize LibreOffice connection
        """
        self.local_context = uno.getComponentContext()
        self.resolver = self.local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", self.local_context)
        self.context = None
        self.desktop = None
        self.document = None

    def connect(self):
        """
        Connect to LibreOffice
        """
        self.context = self.resolver.resolve(
            "uno:socket,host=localhost,port=8100;urp;StarOffice.ComponentContext")
        self.desktop = self.context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.context)
        print("Connected to LibreOffice")

    def open_file(self, file_path):
        """
        Open Excel file
        """
        file_url = uno.systemPathToFileUrl(os.path.abspath(file_path))
        self.document = self.desktop.loadComponentFromURL(
            file_url, "_blank", 0, ())
        print(f"Opened file: {file_path}")

    def save_file(self, file_path):
        """
        Save Excel file
        """
        properties = []
        filter_props = PropertyValue()
        filter_props.Name = "FilterName"
        filter_props.Value = "Calc MS Excel 2007 XML"
        properties.append(filter_props)

        file_url = uno.systemPathToFileUrl(os.path.abspath(file_path))
        self.document.storeAsURL(file_url, tuple(properties))
        print(f"Saved file: {file_path}")

    def close(self):
        """
        Close document and connection
        """
        if self.document:
            self.document.close(True)
        print("Closed LibreOffice connection")

    def convert_xls_to_xlsx(self, input_file, output_file):
        properties = []
        filter_props = PropertyValue()
        filter_props.Name = "FilterName"
        filter_props.Value = "Calc MS Excel 2007 XML"
        properties.append(filter_props)

        input_url = uno.systemPathToFileUrl(os.path.abspath(input_file))
        output_url = uno.systemPathToFileUrl(os.path.abspath(output_file))

        document = self.desktop.loadComponentFromURL(input_url, "_blank", 0, ())
        document.storeToURL(output_url, tuple(properties))
        document.close(True)
        print(f"Converted {input_file} to {output_file}")
        return True

    def convert_xlsx_2_html(self, xlsx_file, html_file, sheet_name):
        """
        Convert an XLSX file to PNG
        """
        input_url = uno.systemPathToFileUrl(os.path.abspath(xlsx_file))
        output_url = uno.systemPathToFileUrl(os.path.abspath(html_file))

        # Open the XLSX file
        self.document = self.desktop.loadComponentFromURL(input_url, "_blank", 0, ())

        # Set export properties for HTML
        properties = []
        filter_props = PropertyValue()
        filter_props.Name = "FilterName"
        filter_props.Value = "HTML (StarCalc)"
        properties.append(filter_props)

        # Export to HTML
        self.document.storeToURL(output_url, tuple(properties))
        print(f"Converted sheet '{sheet_name}' from {xlsx_file} to {html_file}")



def main():
    import argparse
    parser = argparse.ArgumentParser(description="Process excel file")
    parser.add_argument("arg1", help="处理模式")  # convert: 转换xls到xlsx, opensave: 打开保存
    parser.add_argument("arg2", help="文件路径")  # 文件路径
    try:
        parser.add_argument("arg3", help="sheet名称", nargs='?')  # 可选参数
    except:
        pass
    args = parser.parse_args()

    input_file = args.arg2

    if args.arg1 == "convert":  # 执行转换
        if not input_file.endswith(".xls"):
            print("Invalid file format")
            exit(-1)
        ouput_file = args.arg2.replace(".xls", ".xlsx")
        obj = LibreExcel()
        obj.connect()
        obj.convert_xls_to_xlsx(input_file, ouput_file)
        obj.close()
        return 0

    if args.arg1 == "opensave":  # 打开保存，用于公式计算、异常xlsx文件修复等
        obj = LibreExcel()
        obj.connect()
        obj.open_file(input_file)
        obj.save_file(input_file)
        obj.close()
        return 0

    if args.arg1 == "convert_html":  # 执行转换
        if not input_file.endswith(".xlsx"):
            print("Invalid file format")
            exit(-1)
        ouput_file = args.arg2.replace(".xlsx", ".html")
        sheet_name = args.arg3
        obj = LibreExcel()
        obj.connect()
        obj.convert_xlsx_2_html(input_file, ouput_file, sheet_name)
        obj.close()
        return 0
if __name__ == '__main__':
    main()