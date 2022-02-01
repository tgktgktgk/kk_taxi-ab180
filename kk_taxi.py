appkey = '{your_kakao_rest_api_key}'
path = '{folder_where_your_receipts_are_stored}'
savepath = '{xlsx_file_name_with_path}'

"""
terminal에서
pip install opencv-python
pip install openpyxl

*pip 안 되면 pip3
"""

import os
import openpyxl
from openpyxl.styles import Alignment

import cv2
import requests


LIMIT_PX = 3000
LIMIT_BYTE = 3000*3000
LIMIT_BOX = 40


def kakao_ocr_resize(image_path: str):
    """
    ocr detect/recognize api helper
    ocr api의 제약사항이 넘어서는 이미지는 요청 이전에 전처리가 필요.

    pixel 제약사항 초과: resize
    용량 제약사항 초과  : 다른 포맷으로 압축, 이미지 분할 등의 처리 필요. (예제에서 제공하지 않음)

    :param image_path: 이미지파일 경로
    :return:
    """
    image = cv2.imread(image_path)
    height, width, _ = image.shape

    if LIMIT_PX < height or LIMIT_PX < width:
        ratio = float(LIMIT_PX) / max(height, width)
        image = cv2.resize(image, None, fx=ratio, fy=ratio)
        height, width, _ = height, width, _ = image.shape

        # api 사용전에 이미지가 resize된 경우, recognize시 resize된 결과를 사용해야함.
        image_path = "{}_resized.jpg".format(image_path)
        cv2.imwrite(image_path, image)

        return image_path
    return None


def kakao_ocr(image_path: str, appkey: str):
    """
    OCR api request example
    :param image_path: 이미지파일 경로
    :param appkey: 카카오 앱 REST API 키
    """
    API_URL = 'https://dapi.kakao.com/v2/vision/text/ocr'

    headers = {'Authorization': 'KakaoAK {}'.format(appkey)}

    image = cv2.imread(image_path)
    jpeg_image = cv2.imencode(".jpg", image)[1]
    data = jpeg_image.tobytes()


    return requests.post(API_URL, headers=headers, files={"image": data})


def main():
    ## image_path = sys.argv[1]
    image_path = path
    
    # 로컬 디렉토리 안의 파일들
    # children = os.listdir(sys.argv[1])
    children = sorted(os.listdir(image_path), reverse=True)

    # 새로운 Excel 파일 만들기
    wb = openpyxl.Workbook()
    ws = wb.active

    # 헤더 만들기
    ws.append(['', '지출일자', '지출금액', '사용처', '결제수단', '지출목적'])

    # 이미지에서 텍스트 추출하여 값 넣기
    idx = 0
    for i in range(0, len(children)):
        ## image_path = sys.argv[1]
        image_path = path
        image_path = image_path + "/" + children[i]

        if children[i] == '.DS_Store':
            continue

        print("extract text from: {}".format(image_path))

        # 이미지 리사이즈
        resize_impath = kakao_ocr_resize(image_path)
        if resize_impath is not None:
            image_path = resize_impath
            print("원본 대신 리사이즈된 이미지를 사용합니다.")

        # 이미지에서 텍스트 추출
        output = kakao_ocr(image_path, appkey).json()

        vlist = list(output.get('result'))
        
        for j in range(0, len(vlist)):
            if '일시' in str(vlist[j].get('recognition_words')):
                col2 = str(vlist[j+1].get('recognition_words'))[2:10]
            if '금액' in str(vlist[j].get('recognition_words')):
                col3 = str(vlist[j+1].get('recognition_words'))[2:-3]

        idx = idx + 1
        ws.append([idx, col2, col3, '택시', '신용카드', '야근교통비'])

    # 가운데 정렬 후 Excel 파일 저장
    for row in ws.rows:
        for cell in row:
            cell.alignment = Alignment(horizontal='center')

    wb.save(savepath)
    wb.close()

# 실행
main()
