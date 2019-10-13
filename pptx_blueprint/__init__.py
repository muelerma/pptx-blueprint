import pathlib
import pptx
from typing import Union

__Pathlike = Union[str, pathlib.Path]


class Template:
    """Helper class for modifying pptx templates.
    """

    def __init__(self, filename: __Pathlike):
        """Initializes a Template-Modifier.

        Args:
            filename (path-like): file name or path
        """
        self._template_path = filename
        self._presentation = pptx.Presentation(filename)
        pass

    def replace_text(self, label: str, text: str, *, scope=None):
        """Replaces text placeholders on one or many slides.

        Args:
            label (str): label of the placeholder (without curly braces)
            text (str): new content
            scope: None, slide number, Slide object or iterable of Slide objects
        """
        pass

    def replace_picture(self, label: str, filename: __Pathlike):
        """Replaces rectangle placeholders on one or many slides.

        Args:
            label (str): label of the placeholder (without curly braces)
            filename (path-like): path to an image file
        """
        pass

    def replace_table(self, label: str, data, header=False, rownames=False):
        """Replaces rectangle placeholders on one or many slides.

        Args:
            label (str): label of the placeholder (without curly braces)
            data (pandas.DataFrame): table to be inserted into the presentation
        """
        assert isinstance(data, pandas.dataframe)
        
        shapes_to_replace = self._find_shapes(label)

        rows, cols = data.shape
        if header: rows += 1
        if rownames: cols += 1

        for old_shape in shapes_to_replace:
            slide_shapes = old_shape._parent
            table = slide_shapes.add_table(
                rows,
                cols,
                old_shape.left, 
                old_shape.top, 
                old_shape.width,
                old_shape.height
            ).table
            
            # set column widths
            col_width = Length(old_shape.width/cols)
            for i in range(len(table.columns)):
                table.columns[i].width = col_width

            ## fill the table
            for c in range(data.shape[1]):
                for r in range(data.shape[0]):
                    if header and r == 0:
                        c_temp = c
                        ## when rownames, skip first column
                        if rownames: c_temp += 1
                        table.cell(r, c_temp).text = data.columns[c]
                    if rownames and c == 0:
                        r_temp = r
                        ## when header, skip first row
                        if header: r_temp += 1
                        table.cell(r_temp, c).text = str(data.index[r])
                    ## fill table body
                    r_shape = r
                    if header: r_shape += 1
                    c_shape = c
                    if rownames: c_shape += 1
                    table.cell(r_shape, c_shape).text = str(data.iloc[r, c])


    def save(self, filename: __Pathlike):
        """Saves the updated pptx to the specified filepath.

        Args:
            filename (path-like): file name or path
        """
        # TODO: make sure that the user does not override the self._template_path
        pass