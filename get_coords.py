import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, simpledialog

# --- Variáveis globais para armazenar os pontos ---
points = []


def get_mouse_click(event):
    """Callback para capturar o clique do mouse e imprimir as coordenadas."""
    x, y = event.x, event.y
    print(f"Ponto capturado: (x={x}, y={y})")
    points.append((x, y))

    # Se dois pontos foram capturados, imprime a área completa
    if len(points) == 2:
        x0 = min(points[0][0], points[1][0])
        y0 = min(points[0][1], points[1][1])
        x1 = max(points[0][0], points[1][0])
        y1 = max(points[0][1], points[1][1])
        print("\n--- Área de Extração (x0, y0, x1, y1) ---")
        print(f"area = ({x0}, {y0}, {x1}, {y1})")
        print("-------------------------------------------\n")
        points.clear()  # Limpa para a próxima seleção


def main():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal inicial

    # Pede para o usuário selecionar o arquivo PDF
    pdf_path = filedialog.askopenfilename(
        title="Selecione o arquivo PDF", filetypes=[("PDF Files", "*.pdf")])
    if not pdf_path:
        print("Nenhum arquivo selecionado.")
        return

    # Pede o número da página
    page_num = simpledialog.askinteger(
        "Número da Página", "Digite o número da página (começando em 0):", initialvalue=0)
    if page_num is None:
        print("Nenhuma página selecionada.")
        return

    # Renderiza a página do PDF como uma imagem
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_num)
    pix = page.get_pixmap()
    doc.close()

    # Cria uma nova janela para exibir a imagem
    window = tk.Toplevel(root)
    window.title(f"Clique para obter coordenadas - Página {page_num}")

    img_data = pix.tobytes("ppm")
    image = tk.PhotoImage(data=img_data)

    # Cria um canvas para exibir a imagem e capturar cliques
    canvas = tk.Canvas(window, width=pix.width, height=pix.height)
    canvas.create_image(0, 0, anchor="nw", image=image)
    canvas.pack()

    # Associa o evento de clique do mouse à nossa função
    canvas.bind("<Button-1>", get_mouse_click)

    print("Clique no canto SUPERIOR ESQUERDO e depois no canto INFERIOR DIREITO da área desejada.")

    window.mainloop()


if __name__ == "__main__":
    main()
