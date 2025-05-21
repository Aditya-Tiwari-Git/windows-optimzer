#!/usr/bin/env python3
"""
Script to create installer images for NSIS installer
"""

import os
from PIL import Image, ImageDraw, ImageFont

def create_welcome_image(output_path, size=(164, 314), bg_color=(25, 118, 210), text="Windows System Optimizer"):
    """Create a welcome image for the installer"""
    img = Image.new('RGB', size, color=bg_color)
    draw = ImageDraw.Draw(img)
    
    # Try to add some simple text
    try:
        # Try to use a default font
        font_size = 14
        font = ImageFont.truetype("arial.ttf", font_size)
        text_width, text_height = draw.textbbox((0, 0), text, font=font)[2:]
        position = ((size[0] - text_width) / 2, size[1] - 50)
        draw.text(position, text, fill=(255, 255, 255), font=font)
    except Exception as e:
        print(f"Could not add text to welcome image: {e}")
        # Fallback: Just add a simple line
        draw.line([(10, size[1] - 40), (size[0] - 10, size[1] - 40)], fill=(255, 255, 255), width=2)

    # Add a simple gradient
    for y in range(size[1]):
        alpha = 1.0 - (y / size[1]) * 0.7  # Gradient factor
        line_color = (
            int(bg_color[0] * alpha),
            int(bg_color[1] * alpha),
            int(bg_color[2] * alpha)
        )
        draw.line([(0, y), (size[0], y)], fill=line_color)
    
    # Try to add the app icon if available
    try:
        icon_path = os.path.join(os.path.dirname(output_path), "app_icon.png")
        if os.path.exists(icon_path):
            icon = Image.open(icon_path)
            # Resize icon to fit
            icon_size = min(size[0] - 40, 100)
            icon = icon.resize((icon_size, icon_size))
            
            # Calculate position to center the icon
            icon_pos = ((size[0] - icon_size) // 2, 50)
            
            # Paste the icon onto the background
            # If icon has transparency (RGBA), we need to use an alpha mask
            if icon.mode == 'RGBA':
                img.paste(icon, icon_pos, icon)
            else:
                img.paste(icon, icon_pos)
    except Exception as e:
        print(f"Could not add icon to welcome image: {e}")
    
    # Save the image
    img.save(output_path, 'BMP')
    print(f"Created welcome image: {output_path}")

def create_header_image(output_path, size=(150, 57), bg_color=(25, 118, 210), text="Windows System Optimizer"):
    """Create a header image for the installer"""
    img = Image.new('RGB', size, color=bg_color)
    draw = ImageDraw.Draw(img)
    
    # Try to add some simple text
    try:
        # Try to use a default font
        font_size = 10
        font = ImageFont.truetype("arial.ttf", font_size)
        text_width, text_height = draw.textbbox((0, 0), text, font=font)[2:]
        position = ((size[0] - text_width) / 2, (size[1] - text_height) / 2)
        draw.text(position, text, fill=(255, 255, 255), font=font)
    except Exception as e:
        print(f"Could not add text to header image: {e}")
        # Fallback: Just add a simple line
        draw.line([(10, size[1] // 2), (size[0] - 10, size[1] // 2)], fill=(255, 255, 255), width=2)
    
    # Save the image
    img.save(output_path, 'BMP')
    print(f"Created header image: {output_path}")

def main():
    """Main function to create installer images"""
    assets_dir = "assets"
    
    # Ensure the assets directory exists
    os.makedirs(assets_dir, exist_ok=True)
    
    # Create the welcome image
    welcome_path = os.path.join(assets_dir, "installer-welcome.bmp")
    create_welcome_image(welcome_path)
    
    # Create the header image
    header_path = os.path.join(assets_dir, "installer-header.bmp")
    create_header_image(header_path)
    
    print("Installer images created successfully.")

if __name__ == "__main__":
    main() 