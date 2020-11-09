from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from random_user_agent.user_agent import UserAgent
from random_user_agent.params import SoftwareName, OperatingSystem
import time
import json


# Получение фейкового user-agent
def get_agent():
    software_names = [SoftwareName.CHROME.value]
    operating_systems = [OperatingSystem.WINDOWS.value, OperatingSystem.LINUX.value]
    user_agent_rotator = UserAgent(software_names=software_names, operating_systems=operating_systems, limit=100)
    # Получить случайную строку пользовательского агента.
    user_agent = user_agent_rotator.get_random_user_agent()
    return user_agent


chromedriver = '/Users/danil/Downloads/chromedriver'
options = Options()
options.add_argument(f'user-agent={get_agent()}')
driver = webdriver.Chrome(chrome_options=options, executable_path=chromedriver)


dict_profiles = {}  # Словарь вида "имя- [дисциплины]"


# Парсинг сайта
def parsing_data(page: str):
    driver.get(page)
    time.sleep(1)
    profiles = driver.find_elements_by_xpath('// div [@class="nova-v-person-list-item__stack nova-v-person-list-item__'
                                             'stack--gutter-m"]')

    exclusion_disciplien = ['Department of Information Systems', 'Department of Computer Engineering',
                            'Department of Power Engineering', 'Information Technology', 'Civil Engineering',
                            'Department of Applied Mathematics', 'Computer Science', 'Department of Telecommunications',
                            ]
    for profile in profiles:
        name = profile.find_element_by_xpath('.//h5 [@class="nova-e-text nova-e-text--size-l nova-e-text--family-sans-'
                                             'serif nova-e-text--spacing-none nova-e-text--color-inherit nova-e-text--'
                                             'clamp nova-v-person-list-item__title"]').text
        name = name.replace("\'", "\"")
        arr_disciplies = []
        try:
            disciplines = profile.find_elements_by_xpath('.//li[@class="nova-e-list__item nova-v-person-list-item__'
                                                         'info-section-list-item"]//a')
            for discipline in disciplines:
                if (discipline.text in exclusion_disciplien) == True:
                    pass
                else:
                    disp_title = (discipline.text).replace("\'", "\"")
                    arr_disciplies.append(disp_title)
        except:
            pass
        dict_profiles.update(
            {
                name: arr_disciplies
            }
        )


parsing_data('https://www.researchgate.net/institution/Ulyanovsk_State_Technical_University/members')
parsing_data('https://www.researchgate.net/institution/Ulyanovsk_State_Technical_University/members/2')
driver.quit()
with open('profiles.json', 'w') as outfile:
    json.dump(dict_profiles, outfile)
