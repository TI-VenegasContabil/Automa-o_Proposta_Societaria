from src.front_end import Janela, MainFrame


if __name__ == '__main__':

    janela = Janela()

    main_frame = MainFrame(janela)

    main_frame.place(x = 0, y = 0)

    janela.mainloop()