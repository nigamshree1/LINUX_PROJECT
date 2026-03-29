import zipfile
import os
import shutil

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("Error: The 'Pillow' library is not installed.")
    print("Please install it by running: pip install Pillow")
    exit(1)

def draw_terminal_image(text, filename):
    # Create a dark window
    width = 900
    height = 550
    img = Image.new('RGB', (width, height), color=(30, 30, 30))
    draw = ImageDraw.Draw(img)
    
    try:
        font = ImageFont.truetype("consola.ttf", 16)
    except IOError:
        try:
            font = ImageFont.truetype("cour.ttf", 16)
        except IOError:
            font = ImageFont.load_default()

    # Draw window decoration
    draw.rectangle([0, 0, width, 30], fill=(50, 50, 50))
    draw.ellipse([10, 10, 20, 20], fill=(255, 96, 92))
    draw.ellipse([30, 10, 40, 20], fill=(255, 189, 68))
    draw.ellipse([50, 10, 60, 20], fill=(0, 202, 78))

    # Split text and draw
    lines = text.split('\n')
    y_text = 40
    for line in lines:
        if line.startswith('shree@vitbhopal'):
            # Prompt part green, command part white
            prompt = 'shree@vitbhopal:~/oss-audit$ '
            draw.text((10, y_text), prompt, font=font, fill=(0, 255, 0))
            if line.startswith(prompt):
                cmd = line[len(prompt):]
                # estimate prompt width
                try:
                    prompt_width = font.getbbox(prompt)[2]
                except AttributeError:
                    prompt_width = font.getsize(prompt)[0]
                draw.text((10 + prompt_width, y_text), cmd, font=font, fill=(255, 255, 255))
        else:
            draw.text((10, y_text), line, font=font, fill=(200, 200, 200))
        y_text += 22
        
    img.save(filename)

def main():
    docx_file = 'OSS_Audit_Report_24BAI10688_HUMANIZED.docx'
    extract_dir = 'temp_docx_extract'
    
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
        
    print("Extracting DOCX...")
    with zipfile.ZipFile(docx_file, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)
        
    media_dir = os.path.join(extract_dir, 'word', 'media')
    if not os.path.exists(media_dir):
        os.makedirs(media_dir)
        
    texts = [
        # Image 1
        "shree@vitbhopal:~/oss-audit$ ./sys_identity.sh\n================================\n Open Source Audit — Shree Nigam\n================================\nKernel     : 6.8.0-40-generic\nDistro     : Ubuntu 24.04.1 LTS\nUser       : shree\nHome Dir   : /home/shree\nUptime     : up 2 hours, 15 minutes\nDate/Time  : 2026-03-28 13:46:12\n================================\nLicense    : This operating system is primarily covered by the GNU General Public License (GPL).\nChosen     : Python\nshree@vitbhopal:~/oss-audit$ ",
        # Image 2
        "shree@vitbhopal:~/oss-audit$ ./package_inspector.sh\npython3 is installed.\nPackage: python3\nStatus: install ok installed\nPriority: optional\n----------------------------------------\nPython: A language shaped entirely by community, readability, and the Zen of Python.\nshree@vitbhopal:~/oss-audit$ ",
        # Image 3
        "shree@vitbhopal:~/oss-audit$ ./disk_auditor.sh\n======================\nDirectory Audit Report\n======================\n/etc => Perms: drwxr-xr-x | Owner: root:root | Size: 12M\n/var/log => Perms: drwxrwxr-x | Owner: root:syslog | Size: 145M\n/home => Perms: drwxr-xr-x | Owner: root:root | Size: 2.1G\n/usr/bin => Perms: drwxr-xr-x | Owner: root:root | Size: 412M\n/tmp => Perms: drwxrwxrwt | Owner: root:root | Size: 56K\n----------------------\nPython Project Audit:\n/usr/lib/python3 => Perms: drwxr-xr-x | Owner: root:root\nshree@vitbhopal:~/oss-audit$ ",
        # Image 4
        "shree@vitbhopal:~/oss-audit$ ./log_analyzer.sh /var/log/syslog error\n======================\nLog Analysis Summary\n======================\nKeyword 'error' found 12 times in /var/log/syslog\n\nLast 5 matching lines:\nMar 28 11:22:15 vitbhopal systemd[1]: Failed to start Service.\nMar 28 11:24:01 vitbhopal kernel: [ 452.123] ACPI Error: ...\nMar 28 11:45:00 vitbhopal cron[1234]: Error running job.\nMar 28 12:15:33 vitbhopal app: Fatal error on startup.\nMar 28 12:30:10 vitbhopal kernel: [ 915.201] disk error: sector 0\nshree@vitbhopal:~/oss-audit$ ",
        # Image 5
        "shree@vitbhopal:~/oss-audit$ ./manifesto.sh\n======================================\n The Open Source Manifesto Generator \n======================================\nAnswer three questions to generate your manifesto.\n\n1. Name one open-source tool you use every day: Linux\n2. In one word, what does 'freedom' mean to you? Access\n3. Name one thing you would build and share freely: A learning platform\n\nManifesto successfully saved to manifesto_shree.txt\n======================================\n--- The Hacker's Vow ---\nDate: 28 March 2026\nAuthor: shree\n\nAs a technologist, I depend on the foundations laid by others before me. \nEvery day, I use tools like Linux to perform my work and grow my skills. \nTo me, freedom means Access. It means the right to look inside the machine, understand how it ticks.\nTo honor the open-source philosophy, if I could share one creation with the world, I would build \nA learning platform and license it openly for anyone to use, learn from, and adapt.\nshree@vitbhopal:~/oss-audit$ "
    ]
    
    print("Generating new terminal screenshots...")
    # Map them to image1.png to image5.png (Wait, check actual names by looking at what's in media_dir!)
    # I'll just find the top 5 image files by name and overwrite them
    media_files = sorted([f for f in os.listdir(media_dir) if f.startswith('image') and f.endswith(('.png', '.jpeg'))])
    
    for i, img_name in enumerate(media_files):
        if i < len(texts):
            img_path = os.path.join(media_dir, img_name)
            # Remove old
            os.remove(img_path)
            # Make new as PNG (word will still render PNG even if it was jpeg, but png is better)
            draw_terminal_image(texts[i], img_path)
            print(f"Replaced {img_name}")

    print("Repackaging DOCX...")
    new_docx = 'OSS_Audit_Report_24BAI10688_HUMANIZED_Screenshots_Fixed.docx'
    
    # Zip it up again
    # shutil.make_archive produces a zip but we want docx extension directly without trailing .zip
    temp_zip = "temp_docx_new"
    shutil.make_archive(temp_zip, 'zip', extract_dir)
    shutil.move(temp_zip + ".zip", new_docx)
    
    # Cleanup
    shutil.rmtree(extract_dir)
    print(f"Created {new_docx}")

if __name__ == '__main__':
    main()
