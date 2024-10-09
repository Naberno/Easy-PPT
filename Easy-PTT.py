import sys
import os
import time
import threading
import keyboard
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QSystemTrayIcon, QMenu, QAction
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


class MicrophoneControlApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Инициализация атрибутов до вызова initUI()
        self.running = True
        self.mic_enabled = False
        self.ptt_enabled = True  # Флаг для включения/выключения режима PTT
        self.hotkey = 'shift'  # Горячая клавиша для активации микрофона

        # Установка иконки для всего приложения и окна
        self.setWindowIcon(QIcon("icon.png"))

        # Инициализация интерфейса
        self.initUI()

        # Настройка управления микрофоном через pycaw
        devices = AudioUtilities.GetMicrophone()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(interface, POINTER(IAudioEndpointVolume))

        # Поток для мониторинга нажатий
        self.monitor_thread = threading.Thread(target=self.monitor_keys, daemon=True)
        self.monitor_thread.start()

    def initUI(self):
        """Инициализация интерфейса"""
        self.setWindowTitle('Microphone Control')
        self.setGeometry(300, 300, 400, 200)  # Уменьшена высота окна

        # Метка для отображения текущей горячей клавиши
        self.label = QLabel(f'Текущая кнопка активации: {self.hotkey}', self)
        self.label.setGeometry(50, 20, 300, 30)

        # Кнопка для изменения горячей клавиши
        self.change_button = QPushButton('Изменить клавишу', self)
        self.change_button.setGeometry(50, 60, 300, 30)
        self.change_button.clicked.connect(self.change_hotkey)

        # Кнопка включения/выключения режима PTT
        self.toggle_ptt_button = QPushButton('Выключить PTT', self)
        self.toggle_ptt_button.setGeometry(50, 100, 300, 30)
        self.toggle_ptt_button.clicked.connect(self.toggle_ptt)

        # Маленькая кнопка "Выход" в углу экрана
        exit_button = QPushButton('Выход', self)
        exit_button.setGeometry(350, 5, 40, 20)  # Установка кнопки в верхнем правом углу
        exit_button.clicked.connect(self.exit_app)

        # Иконка и меню в трее
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("icon.png"))  # Путь к иконке (замените на путь к .ico или .png)

        # Меню в трее
        self.tray_menu = QMenu()

        # Пункты меню
        self.show_hide_action = QAction("Скрыть", self)  # Изначально текст - "Скрыть"
        self.show_hide_action.triggered.connect(self.show_hide_window)

        toggle_ptt_action = QAction("Выключить PTT", self)
        toggle_ptt_action.triggered.connect(self.toggle_ptt)

        exit_action = QAction("Выход", self)
        exit_action.triggered.connect(self.exit_app)

        # Добавляем пункты в меню
        self.tray_menu.addAction(self.show_hide_action)
        self.tray_menu.addAction(toggle_ptt_action)
        self.tray_menu.addAction(exit_action)

        # Привязываем меню к иконке в трее
        self.tray_icon.setContextMenu(self.tray_menu)

        # Действие при двойном клике на иконку в трее
        self.tray_icon.activated.connect(self.on_tray_icon_click)

        # Показать иконку в трее
        self.tray_icon.show()

    def show_hide_window(self):
        """Скрыть или показать окно, изменяя текст пункта меню"""
        if self.isHidden() or self.isMinimized():
            # Показать окно
            self.show()
            self.activateWindow()
            self.show_hide_action.setText("Скрыть")
        else:
            # Скрыть окно
            self.hide()
            self.show_hide_action.setText("Показать")

    def show_window(self, reason=None, from_menu=False):
        """Показать окно только по двойному клику или через пункт меню "Показать" """
        # Развернуть окно только если вызов пришел из меню (from_menu) или двойным кликом (reason == QSystemTrayIcon.DoubleClick)
        if from_menu or (reason == QSystemTrayIcon.DoubleClick):
            if self.isHidden():
                self.show()
            elif self.isMinimized():
                self.showNormal()
            self.activateWindow()
            # Обновить текст пункта меню
            self.show_hide_action.setText("Скрыть")

    def on_tray_icon_click(self, reason):
        """Обработчик событий для иконки в трее"""
        if reason == QSystemTrayIcon.DoubleClick:
            # Показать окно при двойном клике ЛКМ
            self.show_window(reason=reason)

    def toggle_ptt(self):
        """Включение/выключение режима PTT (Push-to-Talk)"""
        self.ptt_enabled = not self.ptt_enabled
        ptt_status = "Включить PTT" if not self.ptt_enabled else "Выключить PTT"
        self.toggle_ptt_button.setText(ptt_status)  # Обновляем текст кнопки в интерфейсе
        self.tray_icon.contextMenu().actions()[1].setText(ptt_status)  # Обновляем текст кнопки в меню трея

    def change_hotkey(self):
        """Изменение горячей клавиши"""
        # Показать сообщение и ожидать нажатие клавиши
        self.label.setText("Нажмите любую клавишу для установки горячей клавиши...")
        self.repaint()  # Обновить интерфейс

        # Временная остановка потока мониторинга
        self.running = False

        # Ожидание нажатия клавиши
        new_hotkey = keyboard.read_event(suppress=True).name
        self.hotkey = new_hotkey
        self.label.setText(f'Текущая кнопка активации: {self.hotkey}')

        # Возобновить мониторинг нажатий
        self.running = True
        threading.Thread(target=self.monitor_keys, daemon=True).start()

    def monitor_keys(self):
        """Мониторинг нажатий клавиш"""
        while self.running:
            if self.ptt_enabled:
                # Проверяем нажатие только установленной горячей клавиши
                if keyboard.is_pressed(self.hotkey):
                    self.volume.SetMute(0, None)  # Включить микрофон
                else:
                    self.volume.SetMute(1, None)  # Выключить микрофон
            else:
                # Если PTT выключен, микрофон включен всегда
                self.volume.SetMute(0, None)  # Микрофон включен всегда
            time.sleep(0.05)

    def closeEvent(self, event):
        """Переопределение события закрытия окна"""
        event.ignore()
        self.hide()
        self.show_hide_action.setText("Показать")  # Изменить текст пункта меню на "Показать"
        self.tray_icon.showMessage(
            "Microphone Control",
            "Программа свернута в трей. Для выхода используйте контекстное меню в трее.",
            QSystemTrayIcon.Information,
            2000
        )

    def exit_app(self):
        """Функция выхода из программы"""
        self.running = False
        self.tray_icon.hide()
        self.close()
        sys.exit()


def add_to_autostart():
    """Добавление программы в автозагрузку"""
    file_path = os.path.abspath(__file__)
    bat_path = os.path.join(os.getenv('APPDATA'), 'Microsoft\\Windows\\Start Menu\\Programs\\Startup', 'MicrophoneControl.bat')
    with open(bat_path, "w+") as bat_file:
        bat_file.write(f'start "" "{file_path}"')


if __name__ == '__main__':
    app = QApplication(sys.argv)

    # Автозагрузка (раскомментировать при необходимости)
    # add_to_autostart()

    window = MicrophoneControlApp()
    window.show()
    sys.exit(app.exec_())
