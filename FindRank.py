from io import BytesIO
import streamlit as st
from time import sleep
from streamlit_autorefresh import st_autorefresh
from st_aggrid import GridOptionsBuilder, AgGrid
import pandas as pd
import pandas.io.formats.excel

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

st.set_page_config(layout='wide', page_title='Websites Ranks in Google')
st_autorefresh(interval=1e10)

# _________________ Functions _________________
def build_driver(isMobile):
    options = Options()
    prefs = {'profile.default_content_setting_values': {
        'images': 2, 'plugins': 2, 'popups': 2, 'geolocation': 2, 
        'notifications': 2, 'auto_select_certificate': 2,
        'fullscreen': 2, 'mouselock': 2, 'mixed_script': 2, 
        'media_stream': 2, 'media_stream_mic': 2, 'media_stream_camera': 2,
        'protocol_handlers': 2, 'ppapi_broker': 2, 'automatic_downloads': 2, 
        'midi_sysex': 2, 'push_messaging': 2, 'ssl_cert_decisions': 2,
        'metro_switch_to_desktop': 2, 'protected_media_identifier': 2, 
        'app_banner': 2, 'site_engagement': 2, 'durable_storage': 2
        }
        }
    options.add_experimental_option('prefs', prefs)
    if isMobile:
        options.add_argument('--user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 10_3 like Mac OS X) AppleWebKit/602.1.50 (KHTML, like Gecko) CriOS/56.0.2924.75 Mobile/14E5239e Safari/602.1')

    options.add_argument("disable-infobars")
    options.add_argument("--disable-extensions")
    driver = webdriver.Chrome('chromedriver', options=options)
    return driver

# _______ Wait Until Loaded _______
def waitUntilLoaded(driver, findBy, findByValue, allOrOne='one', timeout=1000):
    notLoaded = True
    elapsedTime = 0
    while notLoaded: # wait until page is loaded
        try:
            if allOrOne == 'one':
                element = driver.find_element(findBy, findByValue)
                notLoaded = False
            else:
                element = driver.find_elements(findBy, findByValue)
                if len(element) == 0:
                    _ = 2/0 # make error for going to Exception
                notLoaded = False
        except Exception as err:
            sleep(0.5)
            elapsedTime += 0.5
            if elapsedTime == timeout:
                elapsedTime = 0
                driver.refresh()
    return element

# _____________________________________________________________
def tableStyled(data):
    print(data)
    gb = GridOptionsBuilder.from_dataframe(data)
    gb.configure_pagination(paginationAutoPageSize=True) #Add pagination
    gb.configure_side_bar() #Add a sidebar
    gb.configure_selection('multiple', use_checkbox=True, groupSelectsChildren="Group checkbox select children") 
    
    gridOptions = gb.build()
    
    gridResponse = AgGrid(
        data=data,
        gridOptions=gridOptions,
        theme='streamlit', #Add theme color to the table
        enable_enterprise_modules=True,
        allow_unsafe_jscode=True
    )
    return gridResponse

# _____________________________________________________________
def addSite(siteUrls, _):
    siteUrls = siteUrls.split('\n')
    for siteUrl in siteUrls:
        siteUrl = siteUrl.lower()
        
        if 'www' in siteUrl:
            siteUrl = siteUrl.replace('www.', '').replace('http://', '').replace('https://', '')
        if '.' not in siteUrl:
            st.error('URL is not valid!')
            return -1
        
        siteUrl = siteUrl[0].upper() + siteUrl[1:]
        
        if 'sites' not in st.session_state:
            st.session_state.sites = {}
            st.session_state.sites[siteUrl] = {}
        else:
            if siteUrl not in st.session_state.sites:
                st.session_state.sites[siteUrl] = {}
            else:
                st.warning('This URL has been already added to the list!')
    
# _____________________________________________________________
def addKeyword(keywords, _):
    keywords = keywords.strip().split('\n')
    for keyword in keywords:
        if 'keywords' not in st.session_state:
            st.session_state.keywords = []
            st.session_state.keywords.append(keyword)
        else:
            if keyword not in st.session_state.keywords:
                st.session_state.keywords.append(keyword)
            else:
                st.warning('This keword has been already added to the list!')
                
# _____________________________________________________________
def analyze():
    fullTable = {}
    progressBar = st.progress(0)
    if radioMode == modes[0]:
        isMobile = False
        for siteUrl in st.session_state.sites.keys():
            fullTable[f'{siteUrl}: Rank Number (Desktop)'] = {keyword:-1 for keyword in st.session_state.keywords}
            fullTable[f'{siteUrl}: Found on Page (Desktop)'] = {keyword:-1 for keyword in st.session_state.keywords}
    elif radioMode == modes[1]:
        isMobile = True
        for siteUrl in st.session_state.sites.keys():
            fullTable[f'{siteUrl}: Rank Number (Mobile)'] = {keyword:-1 for keyword in st.session_state.keywords}
            fullTable[f'{siteUrl}: Found on Page (Mobile)'] = {keyword:-1 for keyword in st.session_state.keywords}
    
    lenKeywords = len(st.session_state.keywords)
    
    driver = build_driver(isMobile)
    googleUrl = 'https://www.google.com'
    
    for index, keyword in enumerate(st.session_state.keywords):
        pageNumber = 1
        resultNumber = 1
        isFinished = False
        driver.get(googleUrl)
        
        searchInput = waitUntilLoaded(driver, By.XPATH, '//textarea[@aria-label="Search"]', timeout=5)
        searchInput.send_keys(keyword)
        searchInput.send_keys(Keys.ENTER)
        
        while (not(isFinished)):
            sleep(2)
            if isMobile:
                searchResult = waitUntilLoaded(
                    driver, By.ID, 'rso', timeout=5
                ).find_elements(
                    By.CLASS_NAME, 'pkphOe'
                )
            else:
                searchResult = waitUntilLoaded(
                    driver, By.ID, 'search', timeout=5
                    ).find_elements(
                        By.CLASS_NAME, 'g'
                    )
            for linkContainer in searchResult:
                link = linkContainer.find_element(By.TAG_NAME, 'a')
                link = link.get_attribute('href')
                if 'https' in link or 'http' in link:
                    linkCleaned = link.replace(
                        'https://', ''
                        ).replace(
                            'http://', ''
                            ).replace(
                                'www.', ''
                                ).split('/')[0]

                    # Here we have cleaned link of search result
                    isFirst = True
                    for columnName in fullTable:
                        siteUrl = columnName.split(':')[0]
                        if fullTable[columnName][keyword] == -1:
                            if linkCleaned.lower() == siteUrl.lower():
                                if isFirst:
                                    fullTable[columnName][keyword] = resultNumber
                                    isFirst = False
                                else:
                                    fullTable[columnName][keyword] = pageNumber
                                    isFirst = True
                    resultNumber += 1
                    
            # One page of results is finished reviewing
            for siteUrl in fullTable:
                if fullTable[siteUrl][keyword] == -1:
                    isFinished = False
                    break
            else:
                isFinished = True
            
            if not(isFinished):
                pageNumber += 1
                try:
                    if isMobile:
                        driver.find_element(By.XPATH, f'//a[@aria-label="Next page"]').click()
                    else:
                        driver.find_element(By.XPATH, f'//a[@aria-label="Page {pageNumber}"]').click()
                        
                except Exception as err:
                    for siteUrl in fullTable:
                        if fullTable[siteUrl][keyword] == -1:
                            fullTable[siteUrl][keyword] = '-'

        sleep(5)
        progressBar.progress((index+1)/lenKeywords)

    fullTable['Keywords'] = {keyword:keyword for keyword in st.session_state.keywords}
    st.session_state.fullTable = pd.DataFrame(fullTable)
            
# _____________________________________________________________
# _____________________________________________________________
modes = [
    'Desktop Rank',
    'Mobile Rank',
    ]

buffer = BytesIO()
pandas.io.formats.excel.ExcelFormatter.header_style = None

# ______ Sidebar ______
siteUrls = st.sidebar.text_area(label='Website URLs', placeholder='example.com', help='Add Several URLs with Enter Key')
st.sidebar.button(label='Add URLs', key=1, on_click=addSite, args=(siteUrls, True))

keywords = st.sidebar.text_area(label='Keywords', help='Add Several Keywords with Enter Key')
st.sidebar.button(label='Add Keywords', key=2, on_click=addKeyword, args=(keywords, True))

st.sidebar.markdown('<br>', unsafe_allow_html=True)

radioMode = st.sidebar.radio('Search Modes', options=modes)

st.sidebar.markdown('<br>', unsafe_allow_html=True)

centeralizerCol = st.sidebar.columns([0.3, 0.7 ,0.1])
with centeralizerCol[0]:
    st.write('')
with centeralizerCol[1]:
    st.button(label='Analyse the Result', key='submit', on_click=analyze)
with centeralizerCol[2]:
    st.write('')
    
# _______ Tabs _______
tabs = st.tabs(['Importing Data', 'Result'])
cols = tabs[0].columns(2)
with tabs[0]:
    with cols[0]:
        if 'sites' in st.session_state:
            st.table({'URLs': list(st.session_state.sites.keys())})
        else:
            st.table({'URLs': []})
            
    with cols[1]:
        if 'keywords' in st.session_state:
            st.table({'Keywords': st.session_state.keywords})
        else:
            st.table({'Keywords': []})

with tabs[1]:
    if 'fullTable' in st.session_state:
        tableStyled(st.session_state.fullTable)

        # Save to Excel 
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            fullTable = st.session_state.fullTable.drop('Keywords', axis=1)

            fullTable.to_excel(writer, sheet_name=f'{radioMode}')
            
            workbook  = writer.book
            worksheet = writer.sheets[f'{radioMode}']
            
            cellFormat = workbook.add_format({
                'font_name': 'Helvetica',
                'align': 'center'
                })
            
            headerFormat = workbook.add_format({
                'font_name': 'Helvetica',
                'font_size': 14,
                'bold': True,
                'bg_color': '#D44D5C',
                'color': '#FFFFFF',
                'align': 'center',
                'border': 2
                })
            
            indexFormat = workbook.add_format({
                'font_name': 'Helvetica',
                'font_size': 12,
                'bold': True,
                'bg_color': '#F5E9E2',
                'align': 'center'
                })
            
            for idx, col in enumerate(fullTable):
                series = fullTable[col]
                maxLen = max((
                    series.astype(str).map(len).max(),  # len of largest item
                    len(str(series.name))  # len of column name/header
                    )) + 1
                worksheet.set_column(idx+1, idx+1, width=maxLen, cell_format=cellFormat)

            worksheet.set_column(0, 0, width=10, cell_format=indexFormat)
            worksheet.set_row(0, None, cell_format=headerFormat)
            writer.save()
            
            st.download_button(
                label='Download as Excel File (*.xlsx) \u2B73',
                data=buffer,
                file_name=f'{radioMode}.xlsx',
                mime='application/vnd.ms-excel'
            )