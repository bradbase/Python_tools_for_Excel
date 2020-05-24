# Configuration file for the Sphinx documentation builder.
#
# This file only contains a selection of the most common options. For a full
# list see the documentation:
# http://www.sphinx-doc.org/en/master/config

# -- Path setup --------------------------------------------------------------

# If extensions (or modules to document with autodoc) are in another directory,
# add these directories to sys.path here. If the directory is relative to the
# documentation root, use os.path.abspath to make it absolute, like shown here.
#
import os
import sys
import sphinx

# if not 'READTHEDOCS' in os.environ:
# sys.path.insert(0, os.path.abspath('..'))
sys.path.insert(0, os.path.abspath('.'))
sys.path.append(os.path.abspath('../Python_tools_for_Excel/'))

master_doc = 'index'

# -- Project information -----------------------------------------------------

project = 'Python tools for Excel'
copyright = '2020, Bradley van Ree'
author = 'Bradley van Ree'
version = '0.0.1b0'
# The full version, including alpha/beta/rc tags
release = '0.0.1b0'
# html_logo = 'images/FlyingKoala_ico.svg'


# -- General configuration ---------------------------------------------------

# Add any Sphinx extension module names here, as strings. They can be
# extensions coming with Sphinx (named 'sphinx.ext.*') or your custom
# ones.
extensions = ['recommonmark',
    'sphinx.ext.autodoc',
    'sphinx.ext.todo',
    'sphinx.ext.autosummary',
    'sphinx.ext.viewcode'
]

# Add any paths that contain templates here, relative to this directory.
templates_path = ['_templates']

# List of patterns, relative to source directory, that match files and
# directories to ignore when looking for source files.
# This pattern also affects html_static_path and html_extra_path.
exclude_patterns = ['_build', 'Thumbs.db', '.DS_Store']


# -- Options for HTML output -------------------------------------------------

# The theme to use for HTML and HTML Help pages.  See the documentation for
# a list of builtin themes.
#
# html_theme = 'alabaster'
# html_theme = 'sphinx_rtd_theme'

# Add any paths that contain custom static files (such as style sheets) here,
# relative to this directory. They are copied after the builtin static files,
# so a file named "default.css" will overwrite the builtin "default.css".
html_static_path = []

#---sphinx-themes-----
# html_theme = 'kotti_docs_theme'
# import kotti_docs_theme
# html_theme_path = [kotti_docs_theme.get_theme_dir()]

html_theme = 'msmb_theme'
import msmb_theme
html_theme_path = [msmb_theme.get_html_theme_path()]

# html_theme = 'python_docs_theme'

# html_theme = 'sphinx_pdj_theme'
# import sphinx_pdj_theme
# html_theme_path = [sphinx_pdj_theme.get_html_theme_path()]