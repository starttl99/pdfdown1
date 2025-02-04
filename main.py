#!/usr/bin/env python3
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def automate_opentext_workflow():
    # 1. 다운로드 폴더 설정 (PDF 파일 자동 다운로드용)
    download_dir = os.path.join(os.getcwd(), "downloads")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    # 2. Chrome 옵션 설정 (다운로드 관련 옵션 포함)
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,           # 다운로드 경로 지정
        "download.prompt_for_download": False,                  # 다운로드 시 대화상자 표시 안함
        "plugins.always_open_pdf_externally": True              # PDF를 브라우저 내 열지 않고 바로 다운로드
    }
    chrome_options.add_experimental_option("prefs", prefs)
    # 필요시 headless 모드 추가 가능 (예: chrome_options.add_argument("--headless"))
    
    # 3. WebDriver 객체 생성 (ChromeDriverManager 이용)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(10)  # 암시적 대기 추가 (10초)
    wait = WebDriverWait(driver, 20)

    try:
        # === 1. 한온 해외 Opentext 사이트 접속 및 로그인 자동화 ===

        # 1-1. Opentext 로그인 페이지 접속
        driver.get("https://tradinggrid.opentext.com/tgoportal/#/login")
        
        # 1-2. 로그인 정보 입력 (셀렉터는 실제 사이트에 맞게 수정)
        try:
            username_input = wait.until(EC.presence_of_element_located((By.ID, "userId")))
            password_input = wait.until(EC.presence_of_element_located((By.ID, "password")))
        except TimeoutException:
            print("로그인 페이지 로드 실패: userId 또는 password 입력창을 찾을 수 없습니다.")
            return
        
        username_input.send_keys("youngjang@hyoseong.co.kr")
        password_input.send_keys("Hyoseong02!")
        
        # 1-3. 로그인 버튼 클릭
        try:
            login_button = wait.until(EC.element_to_be_clickable((By.ID, "loginButton")))
            login_button.click()
        except TimeoutException:
            print("로그인 버튼을 찾거나 클릭하는 데 실패했습니다.")
            return
        
        # 1-4. 로그인 후 페이지 전환 대기
        try:
            wait.until(EC.url_contains("tgong"))
            # 페이지 전환 후 안정화를 위해 추가 딜레이
            time.sleep(2)
        except TimeoutException:
            print("로그인 후 Welcome 보드로의 전환이 감지되지 않았습니다.")
            return
        
        # 1-5. Welcome 보드 접속 (필요 시)
        driver.get("https://tradinggrid.opentext.com/tgong/#/homepage")
        time.sleep(2)  # 페이지 로드 대기
        
        # 1-6. Document Manager 접속
        driver.get("https://tradinggrid.opentext.com/#/documents")
        time.sleep(2)
        
        # 1-7. Inbox 클릭
        try:
            inbox_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Inbox']")))
            inbox_element.click()
        except TimeoutException:
            print("Inbox 버튼을 찾거나 클릭하는 데 실패했습니다.")
            return
        
        # === 2. EDI 오더 확인 및 PDF 다운로드 자동화 ===

        # 2-1. 최신 EDI 오더 검색: 'Type' 입력란에 "DELJIT" 입력
        try:
            type_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Type']")))
            type_input.clear()
            type_input.send_keys("DELJIT")
        except TimeoutException:
            print("Type 입력 필드를 찾을 수 없습니다.")
            return

        # 2-2. 'Document ID' 입력란에 "5100001476" 입력
        try:
            doc_id_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Document ID']")))
            doc_id_input.clear()
            doc_id_input.send_keys("5100001476")
        except TimeoutException:
            print("Document ID 입력 필드를 찾을 수 없습니다.")
            return

        # 2-3. 검색 버튼 클릭
        try:
            search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Search')]")))
            search_button.click()
        except TimeoutException:
            print("Search 버튼을 찾거나 클릭하는 데 실패했습니다.")
            return

        # 2-4. 검색 결과에서 해당 오더(문서 ID 5100001476 포함 행) 대기
        try:
            edi_order_row = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//tr[contains(., '5100001476')]")
            ))
        except TimeoutException:
            print("검색 결과에서 해당 오더 행을 찾을 수 없습니다.")
            return

        # 2-5. 해당 오더의 'GENERATE PDF' 버튼 클릭 (오더 행 내부 상대 XPath 사용)
        try:
            generate_pdf_button = edi_order_row.find_element(By.XPATH, ".//button[contains(text(),'GENERATE PDF')]")
            generate_pdf_button.click()
        except NoSuchElementException:
            print("해당 오더 행 내에서 'GENERATE PDF' 버튼을 찾을 수 없습니다.")
            return

        # 2-6. PDF 다운로드 완료 대기 (필요시 다운로드 파일 체크 로직 추가)
        time.sleep(10)
        print("PDF 다운로드가 진행되었습니다. 다운로드 폴더를 확인하세요:", download_dir)

    except Exception as e:
        print("오류 발생:", e)
    finally:
        driver.quit()

if __name__ == "__main__":
    automate_opentext_workflow()
