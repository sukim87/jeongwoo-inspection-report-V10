from io import BytesIO
from pathlib import Path
import re
import tempfile
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageOps

TEMPLATE = Path(__file__).resolve().parents[1] / 'templates' / '검사결과서.xlsx'
PHOTO_SLOTS = [
    ('검사결과서', 'A21'), ('검사결과서', 'G21'),
    ('검사결과서(을지)', 'A11'), ('검사결과서(을지)', 'G11'),
    ('검사결과서(을지)', 'A24'), ('검사결과서(을지)', 'G24'),
    ('검사결과서(을지)', 'A37'), ('검사결과서(을지)', 'G37'),
]
CANVAS_SIZE = (1406, 1000)


def safe(value):
    return re.sub(r'[\\/:*?"<>|]+', '-', str(value or '미입력')).strip()


def prepare_photo(uploaded, degree, fit_mode):
    uploaded.seek(0)
    image = ImageOps.exif_transpose(Image.open(uploaded)).convert('RGB')
    if degree:
        image = image.rotate(-degree, expand=True)

    if fit_mode == '사진 칸 채우기':
        return ImageOps.fit(
            image, CANVAS_SIZE, Image.Resampling.LANCZOS,
            centering=(0.5, 0.5)
        )

    # Default: preserve the full original photo. No cropping.
    contained = ImageOps.contain(image, CANVAS_SIZE, Image.Resampling.LANCZOS)
    canvas = Image.new('RGB', CANVAS_SIZE, 'white')
    left = (CANVAS_SIZE[0] - contained.width) // 2
    top = (CANVAS_SIZE[1] - contained.height) // 2
    canvas.paste(contained, (left, top))
    return canvas


def build_report(data, photos, rotations, fit_modes=None):
    fit_modes = fit_modes or ['전체 사진 표시'] * len(photos)
    wb = load_workbook(TEMPLATE)
    quick = wb['간편입력']
    main = wb['검사결과서']
    extra = wb['검사결과서(을지)']

    quick['A21'] = '검사장소 / Inspection Location'
    quick['B21'] = data.get('inspection_location', '')
    quick_values = {
        'B5': data['customer'], 'B6': data['part_name'],
        'B7': data['order_no'], 'B8': data['po_step_no'],
        'B9': data['date'], 'B10': data['types'],
        'B11': data['spec'], 'B14': data['inspection_content'],
        'B15': data['judgment'], 'B16': '',
        'B17': data['result'], 'B18': data['action'],
        'B19': data['caption'], 'B20': data['inspector'],
    }
    for cell, value in quick_values.items():
        quick[cell] = value
    quick['B9'].number_format = 'yyyy-mm-dd'

    main_values = {
        'J3': data['doc'],
        'C5': data['customer'], 'I5': data['types'],
        'C6': data['part_name'], 'I6': data['po_step_no'],
        'C7': data['order_no'], 'I7': data['date'],
        'A11': data['spec'], 'F11': data['inspection_content'],
        'J11': data.get('inspection_location', ''),
        'C13': data['result'], 'C16': data['action'],
        'A33': data['caption'], 'A37': data['inspector'],
        'J37': data['judgment'],
    }
    for cell, value in main_values.items():
        main[cell] = value
    main['I7'].number_format = 'yyyy-mm-dd'

    extra_values = {
        'J3': data['doc'], 'C5': data['customer'],
        'I5': data['types'], 'C6': data['part_name'],
        'I6': data['po_step_no'], 'C7': data['order_no'],
        'I7': data['date'],
    }
    for cell, value in extra_values.items():
        extra[cell] = value
    extra['I7'].number_format = 'yyyy-mm-dd'

    for ws in (main, extra):
        ws._images = []
    for cell in ('A21', 'G21'):
        main[cell] = ''
    for cell in ('A11', 'G11', 'A24', 'G24', 'A37', 'G37'):
        extra[cell] = ''

    with tempfile.TemporaryDirectory() as temp_dir:
        for index, uploaded in enumerate(photos[:8]):
            degree = rotations[index] if index < len(rotations) else 0
            mode = fit_modes[index] if index < len(fit_modes) else '전체 사진 표시'
            processed = prepare_photo(uploaded, degree, mode)
            path = f'{temp_dir}/{index}.jpg'
            processed.save(path, 'JPEG', quality=92, optimize=True)
            excel_image = XLImage(path)
            excel_image.width = 405
            excel_image.height = 288
            sheet_name, anchor = PHOTO_SLOTS[index]
            wb[sheet_name].add_image(excel_image, anchor)

        output = BytesIO()
        wb.save(output)

    return output.getvalue(), f"{safe(data['doc'])}.xlsx"
