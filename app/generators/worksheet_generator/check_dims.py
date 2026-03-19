
from pptx import Presentation

def check_dimensions(path):
    prs = Presentation(path)
    print(f"Slide Width: {prs.slide_width}")
    print(f"Slide Height: {prs.slide_height}")
    print(f"Width in Inches: {prs.slide_width / 914400.0}")
    print(f"Height in Inches: {prs.slide_height / 914400.0}")

if __name__ == '__main__':
    check_dimensions('worksheet_test.pptx')
