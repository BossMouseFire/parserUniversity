from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from random_user_agent.user_agent import UserAgent
from random_user_agent.params import SoftwareName, OperatingSystem
import time
import json


class Parser:
    # Получение фейкового user-agent
    @staticmethod
    def get_agent():
        software_names = [SoftwareName.CHROME.value]
        operating_systems = [OperatingSystem.WINDOWS.value, OperatingSystem.LINUX.value]
        user_agent_rotator = UserAgent(software_names=software_names, operating_systems=operating_systems, limit=100)
        # Получить случайную строку пользовательского агента.
        user_agent = user_agent_rotator.get_random_user_agent()
        return user_agent

    def get_driver(self):
        chromedriver = './chromedriver'
        options = Options()
        options.headless = True
        options.add_argument(f'user-agent={self.get_agent()}')
        driver = webdriver.Chrome(chrome_options=options, executable_path=chromedriver)
        return driver


class ParserResearchGate(Parser):

    def __init__(self):
        self.dict_href = {} # Словарь вида "имя- [ссылка]"
        self.dict_profiles = {}  # Словарь вида "имя- [дисциплины]"

    def get_url_profiles(self, url):
        driver = self.get_driver()
        driver.get(url)
        time.sleep(1)
        profiles = driver.find_elements_by_xpath(
            '// div [@class="nova-v-person-list-item__stack nova-v-person-list-item__'
            'stack--gutter-m"]')
        for profile in profiles:
            name = profile.find_element_by_xpath(
                './/h5 [@class="nova-e-text nova-e-text--size-l nova-e-text--family-sans-'
                'serif nova-e-text--spacing-none nova-e-text--color-inherit nova-e-text--'
                'clamp nova-v-person-list-item__title"]').text
            name = name.replace("\'", "\"")

            href = profile.find_element_by_xpath(
                './/h5 [@class="nova-e-text nova-e-text--size-l nova-e-text--family-sans-'
                'serif nova-e-text--spacing-none nova-e-text--color-inherit nova-e-text--'
                'clamp nova-v-person-list-item__title"] // a').get_attribute("href")
            href = href.replace("\'", "\"")
            self.dict_href.update(
                {
                    name: href
                }
            )
        driver.quit()

    def update_profile_href(self):
        self.get_url_profiles('https://www.researchgate.net/institution/Ulyanovsk_State_Technical_University/members')
        self.get_url_profiles('https://www.researchgate.net/institution/Ulyanovsk_State_Technical_University/members/2')
        with open('profiles_href.json', 'w') as outfile:
            json.dump(self.dict_href, outfile)

    def get_disciplines_profiles(self, url, name):
        driver = self.get_driver()
        driver.get(url)
        time.sleep(1)
        arr_disciplies = []
        try:
            disciplines = driver.find_elements_by_xpath('// a [@class="nova-e-badge nova-e-badge--color-grey nova-e-'
                                                        'badge--display-inline nova-e-badge--luminosity-medium nova-e-'
                                                        'badge--size-l nova-e-badge--theme-ghost nova-e-badge--radius-'
                                                        'full profile-about__badge"]')
            for discipline in disciplines:
                title = discipline.text.replace("\'", "\"")
                arr_disciplies.append(title)
        except:
            pass

        self.dict_profiles.update(
            {
                name: arr_disciplies
            }
        )
        driver.quit()

    def update_profile_disciplines(self):
        with open("profiles_href.json", 'r', encoding='utf-8') as file:
            disciplines = json.load(file)
            for teacher in disciplines:
                self.get_disciplines_profiles(disciplines[teacher], teacher)
        with open('profiles_disciplines.json', 'w') as outfile:
            json.dump(self.dict_profiles, outfile)

