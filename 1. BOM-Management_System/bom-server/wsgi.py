# PythonAnywhere WSGI 설정 파일
# /var/www/<유저명>_pythonanywhere_com_wsgi.py 에 아래 내용을 붙여넣기

import sys
import os

# 프로젝트 경로 (PythonAnywhere에서 실제 경로로 변경)
project_home = '/home/<유저명>/bom-server'
if project_home not in sys.path:
    sys.path.insert(0, project_home)

os.chdir(project_home)

from app import app as application
