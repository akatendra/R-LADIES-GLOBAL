from datetime import timedelta, datetime
import r_ladies_us_html
import r_ladies_canada_html
import xlsx

from bs4 import BeautifulSoup

# Set up logging
import logging.config

logging.config.fileConfig("logging.ini", disable_existing_loggers=False)
logger = logging.getLogger(__name__)


def prettify_html(html):
    soup = BeautifulSoup(html, 'lxml')
    return soup.prettify()


def parse_letters_links(data):
    soup = BeautifulSoup(data, 'lxml')
    items = soup.select('a[class="cn-char"]')
    logger.debug('###############################################')
    logger.debug(
        f'Number of items founded on page: {len(items)}')
    logger.debug('###############################################')
    letter_links_list = [item['href'] for item in items]
    return letter_links_list


def get_r_ladies(data):
    soup = BeautifulSoup(data, 'lxml')
    items = soup.select('div[class*="cn-list-row"]')
    logger.debug('###############################################')
    logger.debug(
        f'Number of R-Ladies on page: {len(items)}')
    logger.debug('###############################################')
    return items


def reduce_white_spaces(string):
    string = string.replace('\n', '')
    while '  ' in string:
        string = string.replace('  ', ' ')
    return string


websites_max = 0


def parse_r_ladies(r_ladies, xlsx_file_name):
    global websites_max
    for item in r_ladies:
        logger.debug(f'item: {item}')
        r_id = item['id']
        logger.debug(f'id: {r_id}')
        data_entry_id = item['data-entry-id']
        logger.debug(f'data_entry_id: {data_entry_id}')
        data_entry_slug = item['data-entry-slug']
        logger.debug(f'data_entry_slug: {data_entry_slug}')
        image_flag = item.select_one('span[class="cn-image-style"]')
        logger.debug(f'image_flag: {image_flag}')
        if image_flag:
            profile_link = item.select_one('span[class="cn-image-style"] a')[
                'href']
            photo = 'https:' + \
                    item.select_one('span[class="cn-image-style"] img')[
                        'srcset'].split()[0]
        else:
            profile_link = item.select_one('div[class="cn-left"] a')[
                'href']
            photo = ''
        logger.debug(f'profile_link: {profile_link}')
        logger.debug(f'photo: {photo}')

        given_name = item.select_one('span[class="given-name"]').text.strip()
        logger.debug(f'given_name: {given_name}')

        additional_name = item.select_one('span[class="additional-name"]')
        if additional_name:
            additional_name = additional_name.text.strip()
        else:
            additional_name = ''
        logger.debug(f'additional_name: {additional_name}')

        family_name = item.select_one('span[class="family-name"]').text.strip()
        logger.debug(f'family_name: {family_name}')

        honorific_suffix = item.select_one('span[class="honorific-suffix"]')
        logger.debug(f'honorific_suffix_raw: {honorific_suffix}')
        if honorific_suffix:
            honorific_suffix_list = honorific_suffix.text.replace('\"',
                                                                  '').replace(
                '\n', '').strip().split(',')
            logger.debug(f'honorific_suffix_list: {honorific_suffix_list}')
            honorific_suffix_len = len(honorific_suffix_list)
            logger.debug(f'honorific_suffix_len: {honorific_suffix_len}')
            counter = 0
            honorific_suffix = ''
            for suffix in honorific_suffix_list:
                counter += 1
                suffix = suffix.strip()
                if suffix == '':
                    continue
                if counter != honorific_suffix_len:
                    suffix_str = suffix + ', '
                else:
                    suffix_str = suffix
                honorific_suffix += suffix_str
        else:
            honorific_suffix = ''
        logger.debug(f'honorific_suffix: {honorific_suffix}')

        title = item.select_one('span[class="title notranslate"]')
        if title:
            title = title.text.strip()
        else:
            title = ''
        logger.debug(f'title: {title}')

        organization_name = item.select_one(
            'span[class="organization-name notranslate"] a')
        if organization_name:
            organization_name = organization_name.text.strip()
            organization_link = item.select_one(
                'span[class="organization-name notranslate"] a')['href']
        else:
            organization_name = ''
            organization_link = ''
        logger.debug(f'organization_name: {organization_name}')
        logger.debug(f'organization_link: {organization_link}')

        organization_unit = item.select_one(
            'span[class="organization-unit notranslate"]')
        if organization_unit:
            organization_unit = organization_unit.text.strip()
        else:
            organization_unit = ''
        logger.debug(f'organization_unit: {organization_unit}')

        city = item.select_one('span[class="locality"]')
        if city:
            city = city.text.strip()
        else:
            city = ''
        logger.debug(f'city: {city}')

        region = item.select_one('span[class="region"]')
        if region:
            region = region.text.strip()
        else:
            region = ''
        logger.debug(f'region: {region}')

        country_name = item.select_one('span[class="country-name"]')
        if country_name:
            country_name = country_name.text.strip()
        else:
            country_name = ''
        logger.debug(f'country_name: {country_name}')

        websites = item.select('span[class="link cn-link website"]')
        logger.debug(f'websites: {websites}')
        websites_list = []
        for website in websites:
            website = website.find('a')
            if website:
                link_name = website.text.strip()
                website_link = website['href']
            else:
                link_name = ''
                website_link = ''
            logger.debug(f'link_name: {link_name}')
            logger.debug(f'website_link: {website_link}')
            websites_list.append((link_name, website_link))
        logger.debug(f'website_list: {websites_list}')

        website_max_number = len(websites_list)
        if website_max_number > websites_max:
            websites_max = website_max_number

        twitter = item.select_one('a[class="url twitter"]')
        if twitter:
            twitter = twitter['href']
        else:
            twitter = ''
        logger.debug(f'twitter: {twitter}')

        linked_in = item.select_one('a[class="url linked-in"]')
        if linked_in:
            linked_in = linked_in['href']
        else:
            linked_in = ''
        logger.debug(f'linked_in: {linked_in}')

        instagram = item.select_one('a[class="url instagram"]')
        if instagram:
            instagram = instagram['href']
        else:
            instagram = ''
        logger.debug(f'instagram: {instagram}')

        facebook = item.select_one('a[class="url facebook"]')
        if facebook:
            facebook = facebook['href']
        else:
            facebook = ''
        logger.debug(f'facebook: {facebook}')

        # BIO
        bio_list = item.select('li')
        logger.debug(f'bio_list: {bio_list}')
        bio_free = ''
        bio_r_groups = ''
        bio_r_packages = ''
        bio_interests = ''
        bio_contact_method = ''
        for bio in bio_list:
            logger.debug(f'bio raw: {bio}')
            bio = bio.text
            logger.debug(f'bio: {bio}')
            if 'R-Groups' not in bio and 'RGroups' not in bio and 'R-Packages' not in bio and 'RPackages' not in bio and 'Interests' not in bio and 'Contact method' not in bio and 'Contact Method' not in bio:
                bio_free = bio.strip()
                bio_free = reduce_white_spaces(bio_free)
            if 'R-Groups' in bio or 'RGroups' in bio:
                bio_r_groups = bio.replace('R-Groups', '').replace('RGroups',
                                                                   '').replace(
                    ':', '').strip()
                bio_r_groups = reduce_white_spaces(bio_r_groups)
            if 'R-Packages' in bio or 'RPackages' in bio:
                bio_r_packages = bio.replace('R-Packages', '').replace(
                    'RPackages', '').replace(':', '').strip()
                bio_r_packages = reduce_white_spaces(bio_r_packages)
            if 'Interests' in bio:
                bio_interests = bio.replace('Interests', '').replace(':',
                                                                     '').strip()
                bio_interests = reduce_white_spaces(bio_interests)
            if 'Contact method' in bio or 'Contact Method' in bio:
                bio_contact_method_list = bio.replace('Contact method',
                                                      '').replace(
                    'Contact Method',
                    '').replace(':', '').strip().split(',')
                bio_contact_method_list_len = len(bio_contact_method_list)
                counter = 0
                bio_contact_method = ''
                for contact_method in bio_contact_method_list:
                    counter += 1
                    contact_method = contact_method.strip()
                    if counter != bio_contact_method_list_len:
                        contact_method += ', '
                    else:
                        contact_method += ''
                    bio_contact_method += contact_method
        logger.debug(f'bio_free: {bio_free}')
        logger.debug(f'bio_r_groups: {bio_r_groups}')
        logger.debug(f'bio_r_packages: {bio_r_packages}')
        logger.debug(f'bio_interests: {bio_interests}')
        logger.debug(f'bio_contact_method: {bio_contact_method}')

        data_out = {'data_entry_id': data_entry_id,
                    'id': r_id,
                    'data_entry_slug': data_entry_slug,
                    'profile_link': f'=HYPERLINK("{profile_link}", "{profile_link}")',
                    'photo': f'=HYPERLINK("{photo}", "{photo}")',
                    'given_name': given_name,
                    'additional_name': additional_name,
                    'family_name': family_name,
                    'honorific_suffix': honorific_suffix,
                    'title': title,
                    'organization_name': organization_name,
                    'organization_link': f'=HYPERLINK("{organization_link}", "{organization_link}")',
                    'organization_unit': organization_unit,
                    'city': city,
                    'region': region,
                    'country_name': country_name,
                    'twitter': twitter,
                    'linked_in': f'=HYPERLINK("{linked_in}", "{linked_in}")',
                    'instagram': f'=HYPERLINK("{instagram}", "{instagram}")',
                    'facebook': f'=HYPERLINK("{facebook}", "{facebook}")',
                    'bio_r_groups': bio_r_groups,
                    'bio_r_packages': bio_r_packages,
                    'bio_interests': bio_interests,
                    'bio_contact_method': bio_contact_method,
                    'bio_free': bio_free,
                    'website1': '',
                    'website2': '',
                    'website3': '',
                    'website4': '',
                    'website5': ''
                    }

        # Unpack websites
        entries = ('website1', 'website2', 'website3', 'website4', 'website5')
        for entry, website in zip(entries, websites_list):
            data_out[entry] = f'=HYPERLINK("{website[1]}", "{website[0]}")'

        logger.debug(f'data_out: {data_out}')

        xlsx.append_xlsx_file(data_out, xlsx_file_name)

    xlsx.hyperlink_style(xlsx_file_name)
    logger.debug(f'websites_max: {websites_max}')


if __name__ == '__main__':
    html = r_ladies_us_html.html
    r_ladies = get_r_ladies(html)
    parse_r_ladies(r_ladies, 'r_ladies_us.xlsx')
    html = r_ladies_canada_html.html
    r_ladies = get_r_ladies(html)
    parse_r_ladies(r_ladies, 'r_ladies_canada.xlsx')

