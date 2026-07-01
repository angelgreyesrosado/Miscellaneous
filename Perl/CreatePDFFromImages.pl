from PIL import Image

# List of your images in the exact order they appeared
filenames = [
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090421.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090445.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090505.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090517.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090528.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090528.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090539.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090546.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090618.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090627.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090644.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090702.png",
"C:\Users\angel\OneDrive\Pictures\Screenshots\Screenshot_2026-04-04_090724.png"
]

# Open all images and convert to RGB (PDF requirement)
images = [Image.open(f).convert("RGB") for f in filenames]

# Save as a multi-page PDF
output_name = "midsummer_daydream_full_score.pdf"
images[0].save(output_name, save_all=True, append_images=images[1:])

print(f"PDF created: {output_name}")