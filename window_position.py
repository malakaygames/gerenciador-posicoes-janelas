import sys
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QListWidget, QLabel, 
                            QMessageBox, QListWidgetItem)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QColor
import win32gui
import win32con
import win32process
import psutil

# Constantes para estado da janela
RESTORED = 0
MINIMIZED = 1
MAXIMIZED = 2

def get_window_placement(hwnd):
    """Obtém o estado e posição real da janela, mesmo quando maximizada"""
    placement = win32gui.GetWindowPlacement(hwnd)
    # showCmd: 1 = normal, 2 = minimized, 3 = maximized
    is_maximized = placement[1] == 3
    
    if is_maximized:
        # Se está maximizada, retorna a posição normal (não maximizada)
        return {
            'state': MAXIMIZED,
            'rect': placement[4],  # rcNormalPosition
            'current_rect': win32gui.GetWindowRect(hwnd)
        }
    else:
        # Se não está maximizada, retorna a posição atual
        return {
            'state': RESTORED,
            'rect': win32gui.GetWindowRect(hwnd),
            'current_rect': win32gui.GetWindowRect(hwnd)
        }

def set_window_position(hwnd, x, y, width, height, state):
    """Define a posição da janela respeitando seu estado"""
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        is_maximized = placement[1] == 3
        
        # Primeiro, restaura a janela se estiver maximizada
        if is_maximized:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        
        # Define a nova posição
        win32gui.SetWindowPos(
            hwnd, 
            win32con.HWND_TOP,
            x, y, width, height,
            win32con.SWP_SHOWWINDOW
        )
        
        # Se o estado salvo era maximizado, maximiza a janela
        if state == MAXIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    except Exception as e:
        print(f"Erro ao definir posição da janela: {e}")

class WindowManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerenciador de Posições de Janelas")
        self.setGeometry(100, 100, 800, 500)
        
        self.window_positions = self.load_positions()
        self.migrate_data()  # Migra dados antigos para o novo formato
        
        self.active_windows = {}
        
        self.timer = QTimer()
        self.timer.timeout.connect(self.check_windows)
        self.timer.start(1000)
        
        self.init_ui()
    
    def migrate_data(self):
        """Migra dados antigos para o novo formato"""
        migrated = False
        for hwnd, data in self.window_positions.items():
            if 'process_name' not in data:
                data['process_name'] = 'Desconhecido'
                migrated = True
            if 'state' not in data:
                data['state'] = RESTORED
                migrated = True
        
        if migrated:
            self.save_positions()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QHBoxLayout(central_widget)
        
        # Painel esquerdo - Lista de janelas salvas
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        left_layout.addWidget(QLabel("Janelas Salvas:"))
        self.saved_windows_list = QListWidget()
        self.saved_windows_list.itemClicked.connect(self.highlight_saved_window)
        self.update_saved_windows_list()
        left_layout.addWidget(self.saved_windows_list)
        
        # Botões
        buttons_layout = QHBoxLayout()
        remove_btn = QPushButton("Remover")
        remove_btn.clicked.connect(self.remove_window)
        buttons_layout.addWidget(remove_btn)
        
        left_layout.addLayout(buttons_layout)
        
        # Painel direito - Janelas ativas
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        right_layout.addWidget(QLabel("Janelas Ativas:"))
        self.active_windows_list = QListWidget()
        self.active_windows_list.itemClicked.connect(self.highlight_active_window)
        right_layout.addWidget(self.active_windows_list)
        
        save_btn = QPushButton("Salvar Posição da Janela Selecionada")
        save_btn.clicked.connect(self.save_window_position)
        right_layout.addWidget(save_btn)
        
        # Adicionar painéis ao layout principal
        layout.addWidget(left_panel)
        layout.addWidget(right_panel)

    def get_process_name(self, hwnd):
        """Obtém o nome do processo da janela"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            return process.name()
        except:
            return "Desconhecido"

    def get_window_info(self, hwnd):
        """Obtém informações completas da janela"""
        return {
            'title': win32gui.GetWindowText(hwnd),
            'process_name': self.get_process_name(hwnd)
        }

    def window_matches_saved(self, hwnd, saved_info):
        """Verifica se uma janela corresponde às informações salvas"""
        current_info = self.get_window_info(hwnd)
        return (current_info['title'] == saved_info['title'] or 
                (current_info['process_name'] != "Desconhecido" and 
                 current_info['process_name'] == saved_info['process_name']))

    def enum_windows_callback(self, hwnd, windows):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            if not win32gui.GetWindowText(hwnd).startswith("Gerenciador de"):
                window_info = self.get_window_info(hwnd)
                windows.append((hwnd, window_info))

    def get_active_windows(self):
        windows = []
        win32gui.EnumWindows(self.enum_windows_callback, windows)
        return sorted(windows, key=lambda x: x[1]['title'].lower())

    def check_windows(self):
        current_selection = None
        if self.active_windows_list.currentItem():
            current_selection = self.active_windows_list.currentItem().text()
        
        self.active_windows_list.clear()
        self.active_windows.clear()
        
        for hwnd, window_info in self.get_active_windows():
            display_text = f"{window_info['title']} ({window_info['process_name']})"
            self.active_windows[display_text] = hwnd
            item = QListWidgetItem(display_text)
            self.active_windows_list.addItem(item)
            
            if current_selection and display_text == current_selection:
                item.setBackground(QColor('lightblue'))
                self.active_windows_list.setCurrentItem(item)
            
            # Verifica todas as posições salvas
            for saved_hwnd, saved_data in self.window_positions.items():
                if self.window_matches_saved(hwnd, saved_data):
                    try:
                        set_window_position(
                            hwnd,
                            saved_data['x'],
                            saved_data['y'],
                            saved_data['width'],
                            saved_data['height'],
                            saved_data.get('state', RESTORED)
                        )
                    except Exception as e:
                        print(f"Erro ao mover janela {window_info['title']}: {e}")

    def save_window_position(self):
        try:
            selected_item = self.active_windows_list.currentItem()
            if not selected_item:
                QMessageBox.warning(self, "Erro", "Por favor, selecione uma janela ativa primeiro!")
                return
            
            title = selected_item.text()
            if title not in self.active_windows:
                QMessageBox.warning(self, "Erro", "Janela não encontrada na lista de janelas ativas!")
                return
                
            hwnd = self.active_windows[title]
            window_info = self.get_window_info(hwnd)
            
            try:
                placement = get_window_placement(hwnd)
                rect = placement['rect']
                x, y, right, bottom = rect
                width = right - x
                height = bottom - y
                state = placement['state']
            except Exception as e:
                QMessageBox.warning(self, "Erro", f"Não foi possível obter a posição da janela: {str(e)}")
                return
            
            self.window_positions[str(hwnd)] = {
                'title': window_info['title'],
                'process_name': window_info['process_name'],
                'x': x,
                'y': y,
                'width': width,
                'height': height,
                'state': state
            }
            
            self.save_positions()
            self.update_saved_windows_list()
            QMessageBox.information(self, "Sucesso", f"Posição da janela '{window_info['title']}' salva com sucesso!")
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao salvar a posição: {str(e)}")

    def update_saved_windows_list(self):
        current_selection = None
        if self.saved_windows_list.currentItem():
            current_selection = self.saved_windows_list.currentItem().text()
            
        self.saved_windows_list.clear()
        for hwnd, data in self.window_positions.items():
            display_text = f"{data['title']} ({data['process_name']})"
            item = QListWidgetItem(display_text)
            self.saved_windows_list.addItem(item)
            
            if current_selection and display_text == current_selection:
                item.setBackground(QColor('lightblue'))
                self.saved_windows_list.setCurrentItem(item)

    def highlight_active_window(self, item):
        for i in range(self.active_windows_list.count()):
            self.active_windows_list.item(i).setBackground(QColor('white'))
        item.setBackground(QColor('lightblue'))
        
        self.saved_windows_list.clearSelection()
        for i in range(self.saved_windows_list.count()):
            self.saved_windows_list.item(i).setBackground(QColor('white'))

    def highlight_saved_window(self, item):
        for i in range(self.saved_windows_list.count()):
            self.saved_windows_list.item(i).setBackground(QColor('white'))
        item.setBackground(QColor('lightblue'))
        
        self.active_windows_list.clearSelection()
        for i in range(self.active_windows_list.count()):
            self.active_windows_list.item(i).setBackground(QColor('white'))

    def remove_window(self):
        selected_item = self.saved_windows_list.currentItem()
        if not selected_item:
            QMessageBox.warning(self, "Erro", "Selecione uma janela salva primeiro!")
            return
        
        title = selected_item.text()
        hwnd = None
        for saved_hwnd, data in self.window_positions.items():
            display_text = f"{data['title']} ({data['process_name']})"
            if display_text == title:
                hwnd = saved_hwnd
                break
        
        if hwnd and hwnd in self.window_positions:
            del self.window_positions[hwnd]
            self.save_positions()
            self.update_saved_windows_list()
            QMessageBox.information(self, "Sucesso", f"Janela '{title}' removida com sucesso!")

    def load_positions(self):
        try:
            with open('window_positions.json', 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}
    
    def save_positions(self):
        with open('window_positions.json', 'w') as f:
            json.dump(self.window_positions, f, indent=4)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = WindowManager()
    window.show()
    sys.exit(app.exec())
