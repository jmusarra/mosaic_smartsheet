print('TODO: we need a table of fixture types (LED RGBW 16bit) to "Model IDs" (12)')
# Needs to be able to resolve the "Fixture Type" column values into a more generic
#  fixture type that Mosaic can use. This will take some time.

generic_types = {'Conv8': [0, 0, 0],
                'Conv16': [0,1,0],
                'RGB8': [0,5,0],
                'RGB16': [0,5,1],
                'RGBW8': [0, 12, 64],
                'RGBW16': [0, 12, 65]}