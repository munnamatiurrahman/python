import fitz  # PyMuPDF
import os
import sys

def convert_pdfs_here(zoom=2):
    # এই .py ফাইলটা যেই ফোল্ডারে আছে, 👉 সেই একই ফোল্ডারের ভেতরে থাকা সব PDF অটো detect করবে
    base_folder = os.path.dirname(os.path.abspath(sys.argv[0]))

    print(f"📂 Working folder: {base_folder}")

    found = False

    for file in os.listdir(base_folder):
        if file.lower().endswith(".pdf"):
            found = True
            pdf_path = os.path.join(base_folder, file)
            pdf_name = os.path.splitext(file)[0]

            output_folder = os.path.join(base_folder, f"{pdf_name}_images")
            os.makedirs(output_folder, exist_ok=True)

            doc = fitz.open(pdf_path)
            print(f"\n📄 {file} → Total pages: {len(doc)}")

            for i in range(len(doc)):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                img_path = os.path.join(
                    output_folder,
                    f"{pdf_name}_page_{i+1:04d}.png"
                )
                pix.save(img_path)

            doc.close()
            print(f"✅ Images saved in: {output_folder}")

    if not found:
        print("⚠️ এই ফোল্ডারে কোনো PDF পাওয়া যায়নি!")

if __name__ == "__main__":
    convert_pdfs_here(zoom=2)
