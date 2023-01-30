Running from command line
=========================

You can run the command ``pptx-renderer`` from the commandline to convert an
input template into an output presentation.

If you are using `PyGKN <http://docs.gaes.aeroes.internal/PyCouncil/pygkn/>`_
for installing this package, it will be available as a command ``pptx-renderer``
once you have activated the environment.

The following is the syntax of the command.

.. click:: pptx_renderer.command_line:main
    :prog: pptx-renderer
    :nested: full

Example
-------

.. code-block::

    pptx-renderer input_template.pptx output_file.pptx