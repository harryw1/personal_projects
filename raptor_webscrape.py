import csv
import re
from contextlib import closing

from bs4 import BeautifulSoup
from requests import get
from requests.exceptions import RequestException


def simple_get(url):
    """
    Attempts to get the content at 'url' by making
    an HTTP GET request. If the content-type of the
    response is some kind of HTML/XML, return
    text content, otherwise return None.
    """
    try:
        with closing(get(url, stream=True)) as resp:
            if is_good_response(resp):
                return resp.content
            else:
                return None

    except RequestException as e:
        log_error('Error during requests to {0} : {1}'.format(url, str(e)))
        return None


def build_results(url):
    """
    Get the names of all the species and their counts
    from a site for a given day
    """
    webpage = url
    response = simple_get(webpage)

    if response is not None:
        html = BeautifulSoup(response, 'html.parser')
        species_box = html.find('table', attrs={'class': 'splist'})
        species_results = list()
        for t in species_box:
            species_results.append(t.text)

    # Can remove the first item because it holds the text
    # "\nDay's Raptor Counts"
    del species_results[0]

    # Pull the count from the result set
    # using regex to find all numbers on a
    # single line
    raptor_count = list()
    for s in species_results:
        raptor_count.append(re.findall(r'\d+', s))

    # Finally, build a new dictionary
    # containging all of the species
    # and the number of migrants from the day
    results = dict()
    count = 0
    for s in species_results:
        results[s[0:2]] = raptor_count[count]
        count += 1

    return results


def write_to_file(dictionary):
    write = csv.writer(open("output.csv", "w"))
    for key, val in dictionary.items():
        write.writerow([key, val])


def is_good_response(resp):
    """
    Returns true if the response seems to be HTML
    and false otherwise
    """
    content_type = resp.headers['Content-Type'].lower()
    return (resp.status_code == 200
            and content_type is not None
            and content_type.find('html') > -1)


def log_error(e):
    """
    Print errors to the console
    """
    print(e)


def choose_webpage():
    webadress = "https://www.hawkcount.org/day_summary.php?rsite=!!!&rmonth=@@&rday=##&ryear=$$$$"
    site = input('Please enter a site number: ')
    month = input('Please enter a month: ')
    day = input('Please enter a day: ')
    year = input('Please enter a year: ')

    webadress = webadress.replace('!!!', site)
    webadress = webadress.replace('@@', month)
    webadress = webadress.replace('##', day)
    webadress = webadress.replace('$$$$', year)

    return webadress


user_webpage = choose_webpage()
result_set = build_results(user_webpage)
write_to_file(result_set)
