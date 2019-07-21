from pptx import Presentation

prs = Presentation('E:\\test.pptx')
slides = prs.slides

for slide in slides:
    print('silde: %s', slide)
    for shape in slide.shapes:
        if shape.shape_type == 19:
            table = shape
            for row in table.table.rows:
                for cell in row.cells:
                    print(cell.text_frame.text, end='\t')

                print('')