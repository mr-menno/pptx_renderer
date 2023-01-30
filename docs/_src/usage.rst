Using PPTXRenderer
==================

PPTXRenderer allows you to write python code inside a presentation and execute
it like a script. The output is rendered as a new presentation. You can use the
following types of python code inside the input template.

Simple python expressions
-------------------------

You can write simple python expressions inside any text field which gets evaluated
at render time.

For example, you can write ``{{{5.566*3.456}}}`` or ``{{{sum([1123, 123, 123, 4])}}}``


Substituting variable values and function results
-------------------------------------------------

You can write expressions like ``{{{variable_name}}}`` or function call like
``{{{my_function("arg")}}}`` in the presentation. These variables and functions
should be passed in as arguments to the render function.

For example:

If there is a template ppt which contains ``{{{my_variable}}}`` in a text box,
and there is a function call ``{{{calculate_square(my_variable)}}}`` in another
(or same) text box.  You can pass in the input like below.

.. code-block:: python

   from pptx_renderer import PPTXRenderer
   p = PPTXRenderer("template.pptx")

   def sqr(input):
       return input*input

   p.render(
      "output.pptx", 
      {
         "my_variable": 100,
         "calculate_square": sqr
      }
   )

Inserting images
----------------

You can replace shape placeholders with images using ``PPTXRenderer``.
To insert an image, create a rectangle shape which will define where the image
should be placed and the boundaries of the image. Then insert a placeholder
text in the format ``{{{path_to_image:image}}}`` where ``path_to_image`` is a variable
whose value is the path to the image to be inserted.

The renderer will insert the image preserving the aspect ratio within the boundaries
set by the shape.

Inserting tables
----------------

Similar to images, you can replace shape placeholders with tables.
To create a table, create a rectangle shape which will define where the table
should be placed. Then insert a placeholder text in the format 
``{{{table_data:table}}}`` where ``table_data`` is a variable whose value is a
list of lists where the first row is the header of the table and rest is the data.


Writing code in notes
---------------------

Apart from using python expressions inside text boxes, you can write more
elaborate code inside the slide notes surrounded by triple back ticks followed
by ``python`` keyword as shown below.

.. code-block::

   ```python
   # python code
   ```

The code inside this block will get executed before the slide is evaluated.
So, for example, you can define a function side the notes like below.

.. code-block::

   ```python
   def doubler(input):
      return input*input
   ```

Then you can write ``{{{doubler(100)}}}`` inside one of the text boxes in the same
slide or any slide which comes after this slide.