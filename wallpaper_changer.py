import os
import random
import requests
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import time
import ctypes
import win32gui
import win32con
import textwrap

# Constants for wallpaper settings
SPI_SETDESKWALLPAPER = 20
WALLPAPER_FOLDER = r"C:\Wallpapers"  # Folder to save the generated wallpapers
QUOTE_API_URL = "https://zenquotes.io/api/random"
UNSPLASH_API_URL = "https://api.unsplash.com/photos/random"
UNSPLASH_ACCESS_KEY = 'En6dhSAqK9FHK3eb_H3g8iiEeO0p_L5dF3QUfAJSY_c'  # Replace with your Unsplash Access Key
SCREEN_WIDTH = 1920
SCREEN_HEIGHT = 1080

# Ensure the wallpaper folder exists
os.makedirs(WALLPAPER_FOLDER, exist_ok=True)

# Font folder
FONT_FOLDER = r"fonts/"
font_paths = [os.path.join(FONT_FOLDER, font) for font in os.listdir(FONT_FOLDER) if font.endswith(('.ttf', '.otf'))]

# Fetch random quote from ZenQuotes API
def fetch_quote():
    try:
        response = requests.get(QUOTE_API_URL)
        response.raise_for_status()
        data = response.json()
        return data[0]["q"], data[0]["a"]
    except requests.exceptions.RequestException:
        return "No quote available", "ZenQuotes"

# Fetch random image from Unsplash API
def fetch_image():
    try:
        response = requests.get(UNSPLASH_API_URL, params={'client_id': UNSPLASH_ACCESS_KEY, 'query': 'nature', 'orientation': 'landscape'})
        response.raise_for_status()
        img_url = response.json()['urls']['raw']
        img_data = requests.get(img_url).content
        img = Image.open(BytesIO(img_data)).resize((SCREEN_WIDTH, SCREEN_HEIGHT))
        return img
    except requests.exceptions.RequestException:
        return Image.new('RGB', (SCREEN_WIDTH, SCREEN_HEIGHT), color=(0, 0, 0))  # Black image as fallback

# Select a random font from the available fonts
def select_font():
    try:
        font_path = random.choice(font_paths)
        return ImageFont.truetype(font_path, 60)
    except Exception:
        return ImageFont.truetype("arial.ttf", 60)

# Wrap text to fit within the specified width
def wrap_text(text):
    return textwrap.fill(text, width=40)

# Get position for centered text
def get_centered_position(draw, text, font, max_width, max_height, vertical_offset=0):
    wrapped_text = wrap_text(text)
    text_width, text_height = max(draw.textbbox((0, 0), line, font)[2] - draw.textbbox((0, 0), line, font)[0] for line in wrapped_text.split('\n')), \
                             sum(draw.textbbox((0, 0), line, font)[3] - draw.textbbox((0, 0), line, font)[1] for line in wrapped_text.split('\n'))
    x = (max_width - text_width) // 2
    y = (max_height - text_height) // 2 + vertical_offset
    return x, y, wrapped_text

# Draw text with shadow and border effect
def draw_text(draw, text, position, font, text_color, border_color, shadow_offset=2, shadow_color=(0, 0, 0)):
    # Draw the shadow
    shadow_position = (position[0] + shadow_offset, position[1] + shadow_offset)
    draw.text(shadow_position, text, font=font, fill=shadow_color)

    # Draw the border
    border_offset = 7
    for x_offset in [-border_offset, 0, border_offset]:
        for y_offset in [-border_offset, 0, border_offset]:
            if x_offset != 0 or y_offset != 0:
                border_position = (position[0] + x_offset, position[1] + y_offset)
                draw.text(border_position, text, font=font, fill=border_color)

    # Draw the main text
    draw.text(position, text, font=font, fill=text_color)

# Create image with quote
def create_image(quote, author):
    img = fetch_image()
    draw = ImageDraw.Draw(img)
    font = select_font()

    # Combine quote and author
    combined_text = f'"{quote}"\n\n-{author}'
    max_width, max_height = SCREEN_WIDTH - 40, SCREEN_HEIGHT - 40

    # Text and border colors
    text_color = (255, 255, 255)
    border_color = (0, 0, 0)

    # Get centered position for the text
    x, y, wrapped_text = get_centered_position(draw, combined_text, font, max_width, max_height)
    draw_text(draw, wrapped_text, (x, y), font, text_color, border_color)

    # Save the image to the wallpaper folder
    image_path = os.path.join(WALLPAPER_FOLDER, "centered_quote_image_with_shadow_and_border.png")
    img.save(image_path)
    return image_path

# Set the generated image as the wallpaper
def set_wallpaper(image_path):
    ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, image_path, 3)

    # Minimize all open windows to show the desktop
    def show_desktop():
        def enum_windows_callback(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)

        win32gui.EnumWindows(enum_windows_callback, None)

    show_desktop()

# Main loop to update wallpaper
def main():
    while True:
        # print(fetch_quote())
        # exit()
        quote, author = fetch_quote()  # Fetch a new quote
        image_path = create_image(quote, author)  # Create image with the quote
        set_wallpaper(image_path)  # Set the generated image as wallpaper
        # print(f"Updated wallpaper with quote: {quote} — {author}")
        time.sleep(1200)  # Wait for 1 hour before updating the wallpaper again

if __name__ == "__main__":
    main()
