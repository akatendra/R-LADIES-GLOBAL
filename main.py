import time

import request
import r_ladies_us_html

# Set up logging
import logging.config

import scraper

logging.config.fileConfig("logging.ini", disable_existing_loggers=False)
logger = logging.getLogger(__name__)


def spent_time():
    global start_time
    sec_all = time.time() - start_time
    if sec_all > 60:
        minutes = sec_all // 60
        sec = sec_all % 60
        time_str = f'| {int(minutes)} min {round(sec, 1)} sec'
    else:
        time_str = f'| {round(sec_all, 1)} sec'
    start_time = time.time()
    return time_str


def get_html(url):
    html = request.get_request(url)
    html = scraper.prettify_html(html)
    return html


if __name__ == '__main__':
    url_r_ladies_us = 'https://rladies.org/united-states-rladies/'
    url_r_ladies_canada = 'https://rladies.org/canada-rladies'

    time_begin = start_time = time.time()
    # letter_links = get_letter_links(url_r_ladies_us)
    # html = request.get_request(url_r_ladies_us)
    # r_ladies = scraper.get_r_ladies(html)
    # logger.debug(f'R-Ladies: {r_ladies}')
    # html = html_source.html
    # logger.debug(f'html: {html}')
    # html = get_html(url_r_ladies_us)
    # logger.debug(f'{type(html)} | html: {html}')
    # r_ladies = scraper.get_r_ladies(html)
    # data = scraper.parse_r_ladies(r_ladies)
    html = get_html(url_r_ladies_canada)
    logger.debug(f'html: {html}')
