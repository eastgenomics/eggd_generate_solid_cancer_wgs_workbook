from bs4 import BeautifulSoup
import pandas as pd

from bs4 import MarkupResemblesLocatorWarning
import warnings

warnings.filterwarnings("ignore", category=MarkupResemblesLocatorWarning)


def open_html(file: str) -> BeautifulSoup:
    """Open HTML file using BeautifulSoup

    Parameters
    ----------
    file : str
        File path to the HTML

    Returns
    -------
    BeautifulSoup
        BeautifulSoup object for the HTML page
    """

    return BeautifulSoup(file, features="lxml")


def get_images(html: BeautifulSoup) -> list:
    """Get all images in the BeautifulSoup object

    Parameters
    ----------
    html : BeautifulSoup
        BeautifulSoup object

    Returns
    -------
    list
        List of images in the HTML
    """

    return [img for img in html.findAll("img")]


def get_tables(html: str) -> list:
    """Get all the tables in the HTML

    Parameters
    ----------
    html : str
        File path to the HTML file

    Returns
    -------
    list
        List of dataframes for the tables in the HTML file
    """

    return pd.read_html(html)
