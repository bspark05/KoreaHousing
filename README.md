# KoreaHousing

python 한글 issue

파이썬의 기본 인코딩이 'ascii'로 설정되어 있어서 생기는 문제라고 하고,
python 설치 디렉토리 아래에 site-packages/sitecustomize.py 라는 파일을 만들고 아래와 같이 적어주면 해결된다고 하는데,

import sys
reload(sys)
sys.setdefaultencoding("utf-8")
