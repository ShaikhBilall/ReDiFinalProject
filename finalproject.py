from selenium.webdriver.common.by import By
import xlsxwriter

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

def get_video_results(keyword):
    print(f'Processing keyword --->  {keyword}')
    links = {f'https://www.youtube.com/results?search_query={keyword}&sp=EgJAAQ%253D%253D': 'Live',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgJwAQ%253D%253D': '4K',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgIgAQ%253D%253D': 'HD',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgIoAQ%253D%253D': 'Subtitles',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgIwAQ%253D%253D': 'Creative commons',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgJ4AQ%253D%253D': '360',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgPQAQE%253D': 'VR 180',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgI4AQ%253D%253D': '3D',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgPIAQE%253D': 'HDR',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgIYAQ%253D%253D': 'Under 4 minutes',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgIYAw%253D%253D': '4-20 minutes',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgIYAg%253D%253D': 'Over 20 minutes',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgIQAg%253D%253D': 'Channel',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgIQAw%253D%253D': 'Playlist',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgIIAQ%253D%253D': 'Last hour',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgQIAhAB': 'Today',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgQIAxAB': 'This week',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgQIBBAB': 'This month',
             f'https://www.youtube.com/results?search_query={keyword}&sp=EgQIBRAB': 'This year'}

    youtube_data = []
    chrome_options = Options()
    chrome_options.add_argument("--headless")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)



    for i,v in links.items():

        driver.get(i)

        # scrolling to the end of the page
        # https://stackoverflow.com/a/57076690/15164646
        while True:
            # end_result = "No more results" string at the bottom of the page
            # this will be used to break out of the while loop
            end_result =driver.find_elements(by=By.CSS_SELECTOR,value='#message')[0].is_displayed()
            driver.execute_script("var scrollingElement = (document.scrollingElement || document.body);scrollingElement.scrollTop = scrollingElement.scrollHeight;")
            # time.sleep(1) # could be removed
            # print(end_result)

            # once element is located, break out of the loop
            if end_result == True:
                break

        # print('Extracting results. It might take a while...')

        if v == 'Channel':
            for result in driver.find_elements(by=By.ID,value='info-section'):
                title = result.find_element(by=By.CSS_SELECTOR,value='#text.ytd-channel-name').text
                link = result.find_element(by=By.ID,value='main-link').get_attribute(
                    'href')

                # print(title)
                # print(link)

                youtube_data.append({
                    'title': title,
                    'link': link
                })

        elif v == 'Playlist':
            for result in driver.find_elements(by=By.ID,value='#content.ytd-playlist-renderer'):
                title = result.find_element(by=By.CSS_SELECTOR,value='h3.ytd-playlist-renderer').text
                link = result.find_element(by=By.ID,value='#content > a').get_attribute(
                    'href')
                # print(title)
                # print(link)

                youtube_data.append({
                    'title': title,
                    'link': link
                })



        else:
            youtube_data.append({'metric type':v})
            for result in driver.find_elements(by=By.CSS_SELECTOR,value='.text-wrapper.style-scope.ytd-video-renderer'):
                title = result.find_element(by=By.CSS_SELECTOR,value='.title-and-badge.style-scope.ytd-video-renderer').text
                link = result.find_element(by=By.CSS_SELECTOR,value='.title-and-badge.style-scope.ytd-video-renderer a').get_attribute('href')
                channel_name = result.find_element(by=By.CSS_SELECTOR,value='.long-byline').text
                channel_link = result.find_element(by=By.CSS_SELECTOR,value='#text > a').get_attribute('href')
                views = result.find_element(by=By.CSS_SELECTOR,value='.style-scope ytd-video-meta-block').text.split('\n')[0]

                try:
                    time_published = result.find_element(by=By.CSS_SELECTOR,value='.style-scope ytd-video-meta-block').text.split('\n')[1]
                except:
                    time_published = None

                try:
                    snippet = result.find_element(by=By.CSS_SELECTOR,value='.metadata-snippet-container').text
                except:
                    snippet = None

                try:
                    if result.find_element(by=By.CSS_SELECTOR,value='#channel-name .ytd-badge-supported-renderer') is not None:
                        verified_badge = True
                    else:
                        verified_badge = False
                except:
                    verified_badge = None

                try:
                    extensions = result.find_element(by=By.CSS_SELECTOR,value='#badges .ytd-badge-supported-renderer').text
                except:
                    extensions = None
                # print(verified_badge)

                youtube_data.append({
                    'title': title,
                    'link': link,
                    'channel': {'channel_name': channel_name, 'channel_link': channel_link},
                    'views': views,
                    'time_published': time_published,
                    'snippet': snippet,
                    'verified_badge': verified_badge,
                    'extensions': extensions,
                })

        # print(json.dumps(youtube_data, indent=2, ensure_ascii=False))
        # print(youtube_data)

    with xlsxwriter.Workbook(f'{keyword}.xlsx') as workbook:
        worksheet = workbook.add_worksheet()
        # worksheet.write_row(header)
        for row_num, data in enumerate(youtube_data):
            try:
                # print(data)
                data = [data['title'],data['views'],data['channel']['channel_name'],data['snippet'],data['link']]
            except:
                try:
                    data = [data['metric type']]
                except:
                    pass
            worksheet.write_row(row_num, 0, data)
            # print(data)


    driver.quit()

list_of_keywords = ["work","cute","random"]

for keyword in list_of_keywords:
    get_video_results(keyword)
