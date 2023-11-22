from GUI.dash_board import dash_board


if __name__ == '__main__':
    dash_board()

# pyinstaller -F --hidden-import "babel.numbers" Main.py --noconsole