import psutil
import os
import random
import requests
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import sys
import time
import ctypes
import win32com.client
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
if not os.path.exists(WALLPAPER_FOLDER):
    os.makedirs(WALLPAPER_FOLDER)

# Font folder
FONT_FOLDER = r"fonts/"
font_paths = [os.path.join(FONT_FOLDER, font) for font in os.listdir(FONT_FOLDER) if font.endswith(('.ttf', '.otf'))]

def fetch_quote():
    """Fetch random quote from the ZenQuotes API."""
    try:
        response = requests.get(QUOTE_API_URL)
        response.raise_for_status()
        data = response.json()
        return data[0]["q"], data[0]["a"]
    except requests.exceptions.RequestException as e:
        print(f"Error fetching quote: {e}")
        return "No quote available", "ZenQuotes"

def fetch_image():
    """Fetch random image from Unsplash API."""
    try:
        response = requests.get(UNSPLASH_API_URL, params={'client_id': UNSPLASH_ACCESS_KEY, 'query': 'nature', 'orientation': 'landscape'})
        response.raise_for_status()
        data = response.json()
        image_url = data['urls']['raw']

        # Fetch the image from the URL
        image_response = requests.get(image_url)
        image_response.raise_for_status()
        img = Image.open(BytesIO(image_response.content))
        img = img.resize((SCREEN_WIDTH, SCREEN_HEIGHT))
        return img
    except requests.exceptions.RequestException as e:
        print(f"Error fetching image: {e}")
        return Image.new('RGB', (SCREEN_WIDTH, SCREEN_HEIGHT), color=(0, 0, 0))  # Return black background image

def select_font():
    """Select a random font from the available fonts."""
    try:
        font_path = random.choice(font_paths)
        return ImageFont.truetype(font_path, 60)  # Adjust font size as needed
    except Exception as e:
        print(f"Error selecting font: {e}. Using default font.")
        return ImageFont.truetype("arial.ttf", 60)

def wrap_text(draw, text, font, max_width):
    """Wrap text to fit within the specified width."""
    return textwrap.fill(text, width=40)  # Adjust wrap width as needed (usually screen width/character size)

def get_centered_position(draw, text, font, max_width, max_height, vertical_offset=0):
    """Calculate the position to center the text in the image, with an optional vertical offset."""
    wrapped_text = wrap_text(draw, text, font, max_width)
    text_width, text_height = 0, 0
    for line in wrapped_text.split('\n'):
        bbox = draw.textbbox((0, 0), line, font)
        line_width = bbox[2] - bbox[0]
        line_height = bbox[3] - bbox[1]
        text_width = max(text_width, line_width)
        text_height += line_height
    
    x = (max_width - text_width) // 2
    y = (max_height - text_height) // 2 + vertical_offset
    return x, y, wrapped_text

def draw_text_with_border_and_shadow(draw, text, position, font, text_color, border_color, shadow_offset=2, shadow_color=(0, 0, 0)):
    """Draw the text with shadow and border effect."""
    
    # Draw the shadow first, using an offset from the main text
    shadow_position = (position[0] + shadow_offset, position[1] + shadow_offset)
    draw.text(shadow_position, text, font=font, fill=shadow_color)
    
    # Now, draw the border around the text. The border will be drawn in multiple directions.
    border_offset = 4
    for x_offset in [-border_offset, 0, border_offset]:
        for y_offset in [-border_offset, 0, border_offset]:
            if x_offset == 0 and y_offset == 0:
                continue  # Skip the center position; we don't want to draw the border over the main text
            border_position = (position[0] + x_offset, position[1] + y_offset)
            draw.text(border_position, text, font=font, fill=border_color)
    
    # Finally, draw the main text over the shadow and border
    draw.text(position, text, font=font, fill=text_color)


def create_image(quote, author):
    """Create an image with the quote and author attached on top of a random Unsplash image."""
    img = fetch_image()
    draw = ImageDraw.Draw(img)
    font = select_font()

    # Combine quote and author with a separator
    combined_text = f'"{quote}"\n\n- {author}'

    max_width, max_height = SCREEN_WIDTH - 40, SCREEN_HEIGHT - 40

    # Text color and border color
    text_color = (255, 255, 255)  # White text
    border_color = (0, 0, 0)      # Black border (you can adjust this as needed)

    # Draw combined text (quote + author) with shadow, border, and centered
    x, y, wrapped_text = get_centered_position(draw, combined_text, font, max_width, max_height)
    draw_text_with_border_and_shadow(draw, wrapped_text, (x, y), font, text_color, border_color)  # White text with black border

    image_path = os.path.join(WALLPAPER_FOLDER, "centered_quote_image_with_shadow_and_border.png")
    img.save(image_path)
    return image_path

def set_wallpaper(image_path):
    """Set the generated image as the wallpaper."""
    ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, image_path, 3)

def add_to_startup():
    """Add the script to the startup folder."""
    shell = win32com.client.Dispatch("WScript.Shell")
    startup_folder = shell.SpecialFolders("Startup")
    script_name = os.path.join(startup_folder, "SetWallpaper.lnk")
    if not os.path.exists(script_name):
        shortcut = shell.CreateShortCut(script_name)
        shortcut.TargetPath = sys.executable
        shortcut.Arguments = f'"{__file__}"'  # Pass the script itself as argument
        shortcut.save()


def main():


    # Start the new instance
    """Main function to fetch quotes, create wallpaper, and update every hour."""
    while True:
        quote, author = fetch_quote()
        image_path = create_image(quote, author)  # Create image with the quote
        set_wallpaper(image_path)  # Set the generated image as wallpaper
        print(f"Updated wallpaper with quote: {quote} — {author}")
        time.sleep(3600)  # Wait for 1 hour (3600 seconds) before updating the wallpaper again

if __name__ == "__main__":
    add_to_startup()  # Ensure the script runs at startup
    main()  # Start the wallpaper update loop
