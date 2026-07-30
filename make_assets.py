from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
ASSETS.mkdir(exist_ok=True)

size = 256
image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)

# Navy app tile with a warm accent, matching the desktop interface.
draw.rounded_rectangle((8, 8, 248, 248), radius=48, fill="#102A43")
draw.rounded_rectangle((27, 27, 229, 229), radius=36, outline="#294A67", width=4)
draw.ellipse((166, 28, 228, 90), fill="#FF6B3D")

# Document and folded corner.
draw.rounded_rectangle((60, 42, 181, 215), radius=16, fill="#FFFFFF")
draw.polygon(((145, 42), (181, 78), (145, 78)), fill="#DCE8F5")
draw.line((84, 104, 155, 104), fill="#A8BDD0", width=10)
draw.line((84, 130, 145, 130), fill="#A8BDD0", width=10)

# Strong orange confirmation mark.
draw.line((82, 169, 108, 192), fill="#FF6B3D", width=17)
draw.line((106, 192, 158, 142), fill="#FF6B3D", width=17)

output = ASSETS / "wpsfix.ico"
image.save(output, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print(output)
