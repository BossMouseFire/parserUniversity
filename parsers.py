from selenium import webdriver
from random_user_agent.user_agent import UserAgent
from random_user_agent.params import SoftwareName, OperatingSystem
import time
import json
import pytils.translit as tr


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
        moziladriver = './geckodriver'
        options = webdriver.FirefoxOptions()
        options.set_preference('dom.webdriver.enabled', False)
        options.set_preference('general.useragent.override', self.get_agent())
        # options.headless = True
        # options.add_argument(f'user-agent={self.get_agent()}')
        # driver = webdriver.Chrome(chrome_options=options, executable_path=chromedriver)
        driver = webdriver.Firefox(options=options, executable_path=moziladriver)
        return driver


class ParserResearchGate(Parser):

    def __init__(self):
        self.dict_href = {}  # Словарь вида "имя- [ссылка]"
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
        with open('profiles_researchgate.json', 'w') as outfile:
            json.dump(self.dict_profiles, outfile)


class ParserElibrary(Parser):
    def __init__(self):
        self.number = 0
        self.dict_profiles = {
            "Александр": ["fdfd", "fdff"]
        }

    def get_data(self):
        driver = self.get_driver()
        driver.get('https://www.elibrary.ru/authorbox_authors.asp?id=23547')
        # Авторизация на elibrary
        login = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[6]'
                                             '/td/div/table/tbody/tr/td/table[1]/tbody/tr[6]/td/input')
        time.sleep(2)
        login.send_keys('DaniilAbanin')
        password = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[6]/td'
                                                '/div/table/tbody/tr/td/table[1]/tbody/tr[8]/td/input')
        time.sleep(2)
        password.send_keys('qweasdzxc2002Q')
        entry = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[6]/td/'
                                             'div/table/tbody/tr/td/table[1]/tbody/tr[9]/td/input')
        entry.click()
        time.sleep(2)
        group_button = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/'
                                                    'tr[3]/td/form/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/font/a')
        group_button.click()
        number_page = 1
        while number_page != 0:
            table = driver.find_element_by_xpath('//table[@id="restab"]')
            teachers_group = table.find_elements_by_xpath('.//tr')
            for i in range(0, len(teachers_group)):
                try:
                    name = teachers_group[i].find_element_by_xpath('.//td[@align="left"] // font').text
                    button_analytics = teachers_group[i].find_element_by_xpath('.//td[@align="right"] '
                                                                               '// a[@title="Анализ публикационной '
                                                                               'активности автора"]')
                    button_analytics.click()
                    time.sleep(3)
                    buttons = driver.find_elements_by_xpath('// tr[@bgcolor="#f5f5f5"]')
                    button_key = buttons[1].find_element_by_xpath('.// font // a')
                    button_key.click()
                    time.sleep(3)
                    # обработка модального окна
                    driver.switch_to.frame(driver.find_element_by_tag_name("iframe"))
                    array_key = []
                    for key in driver.find_elements_by_xpath('//td[@align="left"] // font'):
                        key = key.text
                        num = ord(key[0])
                        if num >= 65 and num <= 90 or num >= 97 and num <= 122:
                            array_key.append(key.lower())
                    self.dict_profiles.update({
                        name: array_key
                    })
                    print(name, array_key)
                    time.sleep(1)
                    driver.switch_to.default_content()
                    # возврат на страницу с преподавателями
                    driver.back()
                    table = driver.find_element_by_xpath('//table[@id="restab"]')
                    teachers_group = table.find_elements_by_xpath('.//tr')
                    time.sleep(3)
                except Exception as ex:
                    print(ex)

            next_button = driver.find_element_by_xpath('/html/body/div[2]/table/tbody/tr/td/table/tbody'
                                                       '/tr/td[4]/table/tbody/tr[3]/td[2]/a')
            if next_button.text != 'Следующая страница':
                number_page = 0
            else:
                next_button.click()
                time.sleep(3)
        with open('profiles_elibrary.json', 'w', encoding='utf-8') as outfile:
            json.dump(self.dict_profiles, outfile, ensure_ascii=False)

    def add_teachers_in_group(self):
        driver = self.get_driver()
        list_teachers = []
        driver.get('https://www.elibrary.ru/authors.asp?')
        # Вход в аккаунт
        login = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[1]/table/tbody/tr[6]'
                                             '/td/div/div/table[1]/tbody/tr[6]/td/div/input')
        time.sleep(2)
        login.send_keys('DaniilAbanin')
        password = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[1]/table'
                                                '/tbody/tr[6]/td/div/div/table[1]/tbody/tr[8]/td/div/input')
        password.send_keys('qweasdzxc2002Q')
        time.sleep(2)
        entry = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[1]/table/'
                                             'tbody/tr[6]/td/div/div/table[1]/tbody/tr[9]/td/div[2]')
        entry.click()
        time.sleep(2)
        with open("profiles_href.json", 'r', encoding='utf-8') as file:
            disciplines = json.load(file)
            for teacher in disciplines:
                try:
                    time.sleep(3)
                    name_russia = tr.detranslify(teacher)
                    name_russia = name_russia.split(' ')

                    # self.add_teacher_in_group(name_russia[len(name_russia) - 1])

                    # Добавление в список УлГТУ
                    surname = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table'
                                                           '/tbody/tr[2]/td[1]/table/tbody/tr/td/div[1]/div[2]/table[1]'
                                                           '/tbody/tr[2]/td[1]/div/input')
                    surname.clear()
                    time.sleep(3)
                    surname.send_keys(name_russia[len(name_russia) - 1])
                    find_button = driver.find_element_by_xpath(
                        '/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table'
                        '/tbody/tr[2]/td[1]/table/tbody/tr/td/div[1]/div[2]/table[6]'
                        '/tbody/tr[2]/td[6]/div')
                    find_button.click()
                    time.sleep(3)
                    try:
                        highlight_all = driver.find_element_by_xpath(
                            '/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]'
                            '/table/tbody/tr[2]/td[2]/div[2]/div/table/tbody/tr[2]'
                            '/td[2]/a')
                        highlight_all.click()
                        time.sleep(2)
                        add_in_group = driver.find_element_by_xpath(
                            '/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table'
                            '/tbody/tr[2]/td[2]/div[2]/div/table/tbody/tr[5]/td[2]/a')
                        add_in_group.click()
                        time.sleep(2)
                        print(self.number)
                        self.number += 1
                    except:
                        print("Ошибка")
                        list_teachers.append(name_russia[len(name_russia) - 1])
                except:
                    print("Перезагрузка страницы")
                    list_teachers.append(name_russia[len(name_russia) - 1])
                    driver.refresh()
                    time.sleep(10)

        driver.quit()
        return list_teachers


with open("profiles_elibrary.json", 'r', encoding="utf8") as file:
    test1 = json.load(file)

with open("profiles_researchgate.json", 'r', encoding="utf8") as file:
    test2 = json.load(file)