import os
import csv
from PIL import Image, ImageDraw
import zipfile


def load_config(path="config.txt"):
    config = {}
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                if '=' in line:
                    key, value = line.strip().split('=', 1)
                    config[key.strip().lower()] = value.strip()
    else:
        print(f"Config {path} not found. Using defaults.")
    config.setdefault('savepath', 'out')
    config.setdefault('email', '')
    config.setdefault('password', '')
    return config


def load_tasks(path="list.txt"):
    tasks = []
    with open(path, newline='', encoding='utf-8') as f:
        for row in csv.reader(f, delimiter='\t'):
            if len(row) == 3:
                tasks.append(tuple(item.strip() for item in row))
    return tasks


def create_dummy_images(count, folder):
    os.makedirs(folder, exist_ok=True)
    for i in range(1, count + 1):
        img = Image.new('RGB', (200, 200), color='white')
        draw = ImageDraw.Draw(img)
        draw.text((90, 90), str(i), fill='black')
        img.save(os.path.join(folder, f"{i}.png"))


def convert_png_to_jpg(folder):
    for name in sorted(os.listdir(folder)):
        if name.lower().endswith('.png'):
            path = os.path.join(folder, name)
            img = Image.open(path)
            jpg_path = os.path.splitext(path)[0] + '.jpg'
            img.convert('RGB').save(jpg_path, quality=80)
            os.remove(path)


def save_to_pdf(folder, pdf_name):
    images = []
    for name in sorted(os.listdir(folder)):
        if name.lower().endswith('.jpg'):
            img = Image.open(os.path.join(folder, name)).convert('RGB')
            images.append(img)
    if not images:
        print("No JPG files for PDF")
        return
    pdf_path = os.path.join(folder, pdf_name + '.pdf')
    images[0].save(pdf_path, save_all=True, append_images=images[1:])
    for img in images:
        img.close()
    print(f"PDF saved to {pdf_path}")


def save_to_zip(folder, zip_name):
    zip_path = os.path.join(folder, zip_name + '.zip')
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for name in sorted(os.listdir(folder)):
            if name.lower().endswith('.jpg'):
                zf.write(os.path.join(folder, name), arcname=name)
                os.remove(os.path.join(folder, name))
    print(f"ZIP created at {zip_path}")


def main():
    config = load_config()
    tasks = load_tasks()
    for fond, opis, delo in tasks:
        if (fond, opis, delo) != ('48', '1', '3489'):
            continue  # process just one case for demo
        folder_name = f"{fond}-{opis}-{delo}"
        save_folder = os.path.join(config['savepath'], folder_name)
        print(f"Processing {folder_name}...")
        create_dummy_images(3, save_folder)
        convert_png_to_jpg(save_folder)
        save_to_pdf(save_folder, folder_name)
        save_to_zip(save_folder, folder_name)
        print("Done")


if __name__ == '__main__':
    main()
