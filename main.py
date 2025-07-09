import argparse
import win32print
import win32ui
from PIL import Image, ImageWin


def print_image(image_path, printer_name=None, orientation=0,
                paper_width_mm=68, paper_height_mm=130, copies=1,
                scale=100, margin_x=0, margin_y=0, horizontal_offset_mm=0,
                vertical_offset_mm=0, print_width_mm=0):
    if not printer_name:
        printer_name = win32print.GetDefaultPrinter()

    # 打开图片
    image = Image.open(image_path)

    # 旋转图片
    if orientation in [90, 180, 270]:
        image = image.rotate(orientation, expand=True)

    # 创建打印设备上下文
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)

    # 获取打印机DPI（水平和垂直）
    dpi_x = hDC.GetDeviceCaps(88)  # LOGPIXELSX
    dpi_y = hDC.GetDeviceCaps(90)  # LOGPIXELSY

    # 纸张尺寸像素（基于 DPI）
    paper_width_px = int(paper_width_mm / 25.4 * dpi_x)
    paper_height_px = int(paper_height_mm / 25.4 * dpi_y)

    # 边距和偏移像素
    margin_px_x = int(margin_x / 25.4 * dpi_x)
    margin_px_y = int(margin_y / 25.4 * dpi_y)
    offset_px_x = int(horizontal_offset_mm / 25.4 * dpi_x)
    offset_px_y = int(vertical_offset_mm / 25.4 * dpi_y)

    # 计算可用打印宽度（如果print_width_mm <=0则自适应）
    if print_width_mm <= 0:
        print_width_mm = paper_width_mm - 2 * margin_mm - horizontal_offset_mm

    print_width_px = int(print_width_mm / 25.4 * dpi_x)

    # 图片尺寸
    img_w, img_h = image.size

    # 缩放比例（宽度基准）
    scale_ratio = print_width_px / img_w

    # 计算目标打印高度（保持比例）
    print_height_px = int(img_h * scale_ratio)

    # 再应用整体scale百分比
    print_width_px = int(print_width_px * scale / 100)
    print_height_px = int(print_height_px * scale / 100)

    # 绘图左上角位置（考虑边距和偏移）
    draw_x = margin_px_x + offset_px_x
    draw_y = margin_px_y + offset_px_y

    draw_box = (draw_x, draw_y, draw_x + print_width_px, draw_y + print_height_px)

    dib = ImageWin.Dib(image)

    for copy_num in range(copies):
        hDC.StartDoc(f"Print job #{copy_num + 1}")
        hDC.StartPage()
        dib.draw(hDC.GetHandleOutput(), draw_box)
        hDC.EndPage()
        hDC.EndDoc()

    hDC.DeleteDC()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Print image via Windows printer")
    parser.add_argument("image_path", help="Path to the image to print")
    parser.add_argument("--printer-name", default=None, help="Printer name (default printer if omitted)")
    parser.add_argument("--orientation", type=int, default=180, choices=[0, 90, 180, 270], help="Rotation angle")
    parser.add_argument("--paper-width", type=float, default=72, help="Paper width in mm")
    parser.add_argument("--paper-height", type=float, default=130, help="Paper height in mm")
    parser.add_argument("--copies", type=int, default=1, help="Number of copies to print")
    parser.add_argument("--scale", type=int, default=100, help="Scale percentage (100 = no scale)")
    parser.add_argument("--marginx", type=float, default=4, help="Margin in mm")
    parser.add_argument("--marginy", type=float, default=0, help="Margin in mm")
    parser.add_argument("--horizontal-offset", type=float, default=0, help="Horizontal offset in mm")
    parser.add_argument("--vertical-offset", type=float, default=0, help="Vertical offset in mm")
    parser.add_argument("--print-width", type=float, default=68, help="Print width in mm; 0 means auto")

    args = parser.parse_args()

    print_image(
        args.image_path,
        printer_name=args.printer_name,
        orientation=args.orientation,
        paper_width_mm=args.paper_width,
        paper_height_mm=args.paper_height,
        copies=args.copies,
        scale=args.scale,
        margin_x=args.marginx,
        margin_y=args.marginy,
        horizontal_offset_mm=args.horizontal_offset,
        vertical_offset_mm=args.vertical_offset,
        print_width_mm=args.print_width,
    )
