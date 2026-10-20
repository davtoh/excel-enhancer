Excel enhancer
======

Excel enhancer is thought to address the lack of powerful and convenient functions
on Excel in a business or analytical environments. It allows you to import conveniently
made functions on Python to process large amounts of data and even lets you use
your own custom made UDFs as extensions that can be imported. This way the user
can input formulas on cells that otherwise cannot be done with bare excel and
reproduce them later without having to go though the hassle of writing lengthily basic
formulas in the excel syntax.

This packages is a based on xlwings and uses the power of pandas to process table data.

Features :
----------

- Powerful UDFs or custom python functions on excel
- Manage large amount of data with complex operations in a short line or a single UDF.
- Encapsulate different functions into a single one
- Extensions can be imported

--------------

Installation :
--------------

``pip install excel-enhancer``

Quickstart :
------------

Demo code:
'''''''''''''

.. code:: python

    from excel_enhancer import serve
    serve()

--------------

Links :
-------

-  Full Documentation : https://excel-enhancer.readthedocs.io
-  Project home : https://github.com/davtoh/excel-enhancer

.. |Downloads| image:: https://pepy.tech/badge/webbot/week