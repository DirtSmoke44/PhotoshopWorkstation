Как запустить:
Сохрани этот код в файл Photoshop_workstation_python.py в удобной директории.

Для запуска убедись, что установлены все библиотеки

pip install pyinstaller
pip install Pillow
pip install pywin32
pip install PyQt6


pip install PyQt6 pywin32 Pillow - сразу 3 библиотеки устанавливаеты
py -m pip install PyQt6 pywin32 Pillow - если не запускается

pip install PyQt6 PyQt6-Qt6 PyQt6-sip PyQt6-Multimedia Pillow pywin32

!!! pip install pyinstaller PyQt6 PyQt6-Multimedia Pillow pywin32 

Если pywin32 после установки ругается на отсутствие pywintypes — выполни: python -m pywin32_postinstall -install

Проверка, что всё работает:

python -c "import PyQt6; from PyQt6.QtMultimedia import QMediaPlayer; from PIL import Image; import win32com.client; print('OK')"

Запускай простой двойным кликом по .py файлу или через PyCharm/командную строку:

python Photoshop_workstation_python.py

команда для преобразования в 1 exe файл:

ПОКА ИСПОЛЬЗУЕМ ЭТО - pyinstaller --noconsole --onefile --name="Photoshop_Workstation" Photoshop_workstation_python.py

ЭТО ПОКА НЕ ИСПОЛЬЗУЙ! - pyinstaller --noconfirm --clean --onefile --windowed --name "PhotoshopWorkStationPlus" --hidden-import=win32com.client --hidden-import=pythoncom --hidden-import=pywintypes --collect-submodules win32com --collect-all PyQt6 Photoshop_workstation_python.py

ЗАПУСК программы: py Photoshop_workstation_python.py

