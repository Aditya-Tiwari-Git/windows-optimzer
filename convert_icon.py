"""
Convert PNG icon to ICO format for better Windows application integration.
"""

from PIL import Image
import os

def png_to_ico(png_file, ico_file):
    """Convert PNG file to ICO format."""
    try:
        # Open the PNG image
        img = Image.open(png_file)

        # Convert to ICO format with multiple sizes
        # Windows typically uses these sizes: 16, 32, 48, 64, 128, 256
        sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
        img.save(ico_file, sizes=sizes)
        print(f"Successfully converted {png_file} to {ico_file}")
        return True
    except Exception as e:
        print(f"Error converting icon: {str(e)}")
        return False

if __name__ == "__main__":
    # Input and output files
    png_file = "app_icon.png"
    ico_file = "assets/app_icon.ico"
    
    # Make sure assets directory exists
    os.makedirs("assets", exist_ok=True)
    
    # Convert PNG to ICO
    png_to_ico(png_file, ico_file) 