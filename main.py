import win32print
import win32ui
import win32con
import win32api
from PIL import Image
import sys
import os
import argparse
from datetime import datetime

def print_image(image_path, printer_name=None, orientation=0,
                paper_width=76, paper_height=130, copies=1,
                scale=100, margin=0, horizontal_offset=0,
                vertical_offset=0, print_width=None):
    """
    增强版图片打印函数，支持自定义纸张尺寸和方向

    参数:
    image_path: 图片文件路径
    printer_name: 打印机名称(默认使用默认打印机)
    orientation: 打印方向(0, 90, 180, 270)
    paper_width: 纸张宽度(毫米)
    paper_height: 纸张高度(毫米)
    copies: 打印份数
    scale: 缩放比例(百分比)
    margin: 页边距(毫米)
    horizontal_offset: 水平偏移(毫米)
    vertical_offset: 垂直偏移(毫米)
    print_width: 打印宽度(毫米)，None表示使用纸张宽度
    """

    # 打开图片并获取原始DPI
    try:
        img = Image.open(image_path)
        dpi = img.info.get('dpi', (72, 72))  # 默认使用72 DPI
        img.close()
    except Exception as e:
        print(f"无法打开图片: {e}")
        return False

    # 获取打印机句柄
    if not printer_name:
        printer_name = win32print.GetDefaultPrinter()

    try:
        hprinter = win32print.OpenPrinter(printer_name)
    except Exception as e:
        print(f"无法打开打印机: {e}")
        return False

    # 获取打印机默认设置
    devmode = win32print.GetPrinter(hprinter, 2)["pDevMode"]

    # 设置打印方向 (0, 90, 180, 270)
    if orientation in [0, 180]:
        devmode.Orientation = win32con.DMORIENT_PORTRAIT
    else:
        devmode.Orientation = win32con.DMORIENT_LANDSCAPE

    # 设置自定义纸张尺寸（毫米转换为0.1毫米单位）
    devmode.PaperSize = win32con.DMPAPER_USER
    devmode.PaperWidth = int(paper_width * 10)  # 转换为0.1毫米
    devmode.PaperLength = int(paper_height * 10)  # 转换为0.1毫米

    # 设置字段标志，通知系统我们修改了纸张尺寸
    devmode.Fields |= win32con.DM_PAPERSIZE
    devmode.Fields |= win32con.DM_PAPERWIDTH
    devmode.Fields |= win32con.DM_PAPERLENGTH

    # 设置打印份数
    devmode.Copies = copies

    # 创建打印机DC
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)
    hdc.SetMapMode(win32con.MM_TWIPS)  # 使用1/1440英寸为单位

    # 开始打印作业
    job_name = f"Python Print Job - {os.path.basename(image_path)}"
    hdc.StartDoc(job_name)
    hdc.StartPage()

    # 计算实际打印宽度
    if print_width is None:
        print_width = paper_width

    # 计算页面尺寸和边距（转换为twips）
    mm_to_twips = 56.7  # 1毫米 = 56.7 twips
    paper_width_twips = int(paper_width * mm_to_twips)
    paper_height_twips = int(paper_height * mm_to_twips)
    margin_twips = int(margin * mm_to_twips)
    horizontal_offset_twips = int(horizontal_offset * mm_to_twips)
    vertical_offset_twips = int(vertical_offset * mm_to_twips)

    # 加载图片到内存DC
    try:
        img = Image.open(image_path)

        # 根据方向旋转图像
        if orientation == 90:
            img = img.rotate(270, expand=True)
        elif orientation == 180:
            img = img.rotate(180, expand=True)
        elif orientation == 270:
            img = img.rotate(90, expand=True)

        dib = win32ui.CreateBitmap()
        dib.CreateCompatibleBitmap(hdc, img.width, img.height)
        mem_dc = hdc.CreateCompatibleDC()
        mem_dc.SelectObject(dib)

        # 将图片数据复制到内存DC
        for y in range(img.height):
            for x in range(img.width):
                color = img.getpixel((x, y))
                # 处理RGBA图像
                if isinstance(color, tuple) and len(color) > 3:
                    color = color[:3]
                mem_dc.SetPixel(x, y, win32api.RGB(*color))
    except Exception as e:
        print(f"处理图片时出错: {e}")
        hdc.EndDoc()
        return False

    # 计算缩放比例
    scale_factor = scale / 100.0

    # 计算图片尺寸（考虑DPI）
    img_width_inch = img.width / dpi[0]
    img_height_inch = img.height / dpi[1]

    # 转换为设备单位（twips）
    img_width_twips = int(img_width_inch * 1440 * scale_factor)
    img_height_twips = int(img_height_inch * 1440 * scale_factor)

    # 限制打印宽度
    max_print_width_twips = int(print_width * mm_to_twips)
    if img_width_twips > max_print_width_twips:
        ratio = max_print_width_twips / img_width_twips
        img_width_twips = max_print_width_twips
        img_height_twips = int(img_height_twips * ratio)

    # 计算位置（考虑偏移和边距）
    x_pos = margin_twips + horizontal_offset_twips
    y_pos = margin_twips + vertical_offset_twips

    # 绘制图片
    hdc.StretchBlt(
        x_pos, y_pos,
        img_width_twips, img_height_twips,
        mem_dc,
        0, 0, img.width, img.height,
        win32con.SRCCOPY
    )

    # 结束打印作业
    hdc.EndPage()
    hdc.EndDoc()

    # 清理资源
    del mem_dc
    del dib
    img.close()
    hdc.DeleteDC()

    print(f"成功发送打印作业到 {printer_name}")
    return True

def list_printers():
    """列出所有可用的打印机"""
    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )
    print("\n可用的打印机:")
    for i, printer in enumerate(printers):
        print(f"{i+1}. {printer[2]}")

def main():
    parser = argparse.ArgumentParser(
        description='Xprinter XP-D10 图片打印中间件',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('image_path', help='要打印的图片文件路径')
    parser.add_argument('-p', '--printer', help='指定打印机名称')
    parser.add_argument('-o', '--orientation', type=int, choices=[0, 90, 180, 270],
                        default=0, help='打印方向 (0, 90, 180, 270)')
    parser.add_argument('-pw', '--paper_width', type=float, default=76.0,
                        help='纸张宽度(毫米) (默认: 76.0)')
    parser.add_argument('-ph', '--paper_height', type=float, default=130.0,
                        help='纸张高度(毫米) (默认: 130.0)')
    parser.add_argument('-c', '--copies', type=int, default=1,
                        help='打印份数 (默认: 1)')
    parser.add_argument('-sc', '--scale', type=int, default=100,
                        help='缩放比例 (百分比, 默认: 100)')
    parser.add_argument('-m', '--margin', type=float, default=0.0,
                        help='页边距 (毫米, 默认: 0)')
    parser.add_argument('-ho', '--horizontal_offset', type=float, default=0.0,
                        help='水平偏移 (毫米, 默认: 0)')
    parser.add_argument('-vo', '--vertical_offset', type=float, default=0.0,
                        help='垂直偏移 (毫米, 默认: 0)')
    parser.add_argument('-w', '--print_width', type=float,
                        help='打印宽度(毫米) (默认使用纸张宽度)')
    parser.add_argument('-l', '--list', action='store_true',
                        help='列出所有可用的打印机')

    args = parser.parse_args()

    if args.list:
        list_printers()
        return

    if not os.path.exists(args.image_path):
        print(f"错误: 文件 '{args.image_path}' 不存在")
        return

    print("=" * 70)
    print(f"Xprinter XP-D10 图片打印中间件 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    print(f"图片: {args.image_path}")
    print(f"打印机: {args.printer if args.printer else '默认打印机'}")
    print(f"方向: {args.orientation}°")
    print(f"纸张尺寸: {args.paper_width}mm x {args.paper_height}mm")
    print(f"打印宽度: {args.print_width if args.print_width else args.paper_width}mm")
    print(f"份数: {args.copies}")
    print(f"缩放: {args.scale}%")
    print(f"边距: {args.margin}mm")
    print(f"水平偏移: {args.horizontal_offset}mm")
    print(f"垂直偏移: {args.vertical_offset}mm")
    print("-" * 70)

    success = print_image(
        args.image_path,
        printer_name=args.printer,
        orientation=args.orientation,
        paper_width=args.paper_width,
        paper_height=args.paper_height,
        copies=args.copies,
        scale=args.scale,
        margin=args.margin,
        horizontal_offset=args.horizontal_offset,
        vertical_offset=args.vertical_offset,
        print_width=args.print_width
    )

    if success:
        print("打印作业已成功提交!")
    else:
        print("打印过程中出现错误")

if __name__ == "__main__":
    main()