import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import requests

def create_driver():
    """WebDriver 설정 및 생성"""
    download_dir = os.path.join(os.getcwd(), "downloads")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.set_capability(
        "goog:loggingPrefs", {"performance": "ALL", "browser": "ALL"}
    )
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(20)
    
    return driver, download_dir

def find_and_click_element(driver, wait, selectors, scroll=True):
    """요소를 찾아서 클릭하는 함수"""
    for selector in selectors:
        try:
            print(f"선택자 시도 중: {selector}")
            element = wait.until(EC.element_to_be_clickable(selector))
            if element:
                if scroll:
                    try:
                        driver.execute_script("arguments[0].scrollIntoView(true);", element)
                        time.sleep(1)
                    except Exception as e:
                        print(f"스크롤 중 오류 발생 (무시하고 계속): {str(e)}")
                
                # 여러 클릭 방법 시도
                try:
                    # 1. JavaScript 클릭
                    driver.execute_script("arguments[0].click();", element)
                except Exception:
                    try:
                        # 2. 일반 클릭
                        element.click()
                    except Exception:
                        try:
                            # 3. ActionChains 사용
                            ActionChains(driver).move_to_element(element).click().perform()
                        except Exception as e:
                            print(f"클릭 시도 실패: {str(e)}")
                            continue
                
                print(f"요소를 성공적으로 찾아 클릭했습니다: {selector}")
                time.sleep(2)
                return element
        except Exception as e:
            print(f"선택자 {selector} 시도 실패: {str(e)}")
            continue
    return None

def wait_for_angular_load(driver, wait):
    """Angular 앱이 로드될 때까지 대기"""
    try:
        wait.until(lambda driver: driver.execute_script(
            'return window.getAllAngularTestabilities().findIndex(x => !x.isStable()) === -1'))
        return True
    except Exception as e:
        print(f"Angular 로드 대기 중 오류: {str(e)}")
        return False

def automate_opentext_workflow():
    driver = None
    try:
        # WebDriver 설정 및 생성
        driver, download_dir = create_driver()
        wait = WebDriverWait(driver, 45)

        # 1. 로그인 페이지 접속
        driver.get("https://tradinggrid.opentext.com/tgoportal/#/login")
        print("로그인 페이지 접속 완료")
        
        # Angular 앱 로드 대기
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-root")))
        time.sleep(7)
        
        # 2. 로그인 폼 요소 찾기
        input_selectors = [
            (By.XPATH, "//input[@type='text' or @type='email']"),
            (By.XPATH, "//input[contains(@class, 'email') or contains(@class, 'username')]"),
            (By.CSS_SELECTOR, "input[formcontrolname='userId']"),
            (By.ID, "userId")
        ]
        
        username_input = None
        for selector in input_selectors:
            try:
                username_input = wait.until(EC.presence_of_element_located(selector))
                if username_input:
                    break
            except TimeoutException:
                continue

        password_selectors = [
            (By.XPATH, "//input[@type='password']"),
            (By.CSS_SELECTOR, "input[formcontrolname='password']"),
            (By.ID, "password")
        ]
        
        password_input = None
        for selector in password_selectors:
            try:
                password_input = wait.until(EC.presence_of_element_located(selector))
                if password_input:
                    break
            except TimeoutException:
                continue

        if not username_input or not password_input:
            print("로그인 폼 요소를 찾을 수 없습니다.")
            return

        # 3. 로그인 정보 입력
        username_input.clear()
        username_input.send_keys("youngjang@hyoseong.co.kr")
        time.sleep(1)
        password_input.clear()
        password_input.send_keys("Hyoseong02!")
        time.sleep(1)

        # Sign in 버튼 찾기 및 클릭
        sign_in_selectors = [
            (By.XPATH, "//button[contains(text(), 'Sign in')]"),
            (By.XPATH, "//button[text()='Sign in']"),
            (By.XPATH, "//*[text()='Sign in']"),
            (By.CSS_SELECTOR, "button.signin-button"),
            (By.CSS_SELECTOR, "button[type='submit']")
        ]

        sign_in_button = find_and_click_element(driver, wait, sign_in_selectors)
        if not sign_in_button:
            print("Sign in 버튼을 찾을 수 없습니다.")
            return

        print("로그인 시도 완료")
        time.sleep(10)

    finally:
        if driver:
            driver.quit()

def navigate_to_inbox(driver, wait):
    """Document Manager로 이동 후 Inbox 클릭"""
    try:
        print("Document Manager로 이동 시도 중...")
        
        # 명확한 Document Manager 요소 찾기
        doc_manager_selectors = [
            (By.CSS_SELECTOR, "div.portal-tile-title div.portal-title-header-text"),
            (By.XPATH, "//div[contains(@class, 'portal-title-header-text') and contains(text(), 'Document Manager')]"),
            (By.XPATH, "//div[contains(@class, 'portal-tile-title')]//div[contains(text(), 'Document Manager')]")
        ]
        
        # Document Manager 클릭
        doc_manager = None
        for selector in doc_manager_selectors:
            try:
                doc_manager = wait.until(EC.element_to_be_clickable(selector))
                if doc_manager:
                    # 요소가 클릭 가능할 때까지 짧게 대기
                    time.sleep(2)
                    
                    try:
                        # JavaScript로 클릭 시도
                        driver.execute_script("arguments[0].click();", doc_manager)
                    except:
                        try:
                            # 일반 클릭 시도
                            doc_manager.click()
                        except:
                            # ActionChains 사용 시도
                            ActionChains(driver).move_to_element(doc_manager).click().perform()
                    
                    print("Document Manager 클릭 성공")
                    break
            except Exception as e:
                print(f"선택자 {selector} 실패: {str(e)}")
                continue
                
        if not doc_manager:
            print("Document Manager 요소를 찾을 수 없습니다")
            return False
            
        # Document Manager 페이지 로드 대기
        time.sleep(5)
        
        # Inbox 클릭
        inbox_selector = (By.CSS_SELECTOR, "li#dm_dibHeader.dm-tile-header")
        inbox_element = wait.until(EC.element_to_be_clickable(inbox_selector))
        
        if inbox_element:
            try:
                driver.execute_script("arguments[0].click();", inbox_element)
                print("Inbox 클릭 성공")
                time.sleep(5)
                return True
            except Exception as e:
                print(f"Inbox 클릭 실패: {str(e)}")
                return False
                
        return False
        
    except Exception as e:
        print(f"Document Manager 네비게이션 오류: {str(e)}")
        return False

def search_and_download(driver, wait):
    """검색 및 PDF 다운로드 수행"""
    try:
        # Type 입력
        type_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Type']")))
        type_input.clear()
        type_input.send_keys("DELJIT")
        time.sleep(1)

        # Document ID 입력
        doc_id_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Document ID']")))
        doc_id_input.clear()
        doc_id_input.send_keys("5100001476")
        time.sleep(1)

        # 검색 버튼 클릭
        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Search')]")))
        if not find_and_click_element(driver, wait, [(By.XPATH, "//button[contains(text(),'Search')]")]):
            return False

        time.sleep(5)
        
        # 검색 결과 처리
        try:
            row = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[contains(., '5100001476')]")))
            pdf_button = row.find_element(By.XPATH, ".//button[contains(text(),'GENERATE PDF')]")
            
            if find_and_click_element(driver, wait, [(By.XPATH, ".//button[contains(text(),'GENERATE PDF')]")]):
                print("PDF 생성 버튼 클릭 성공")
                time.sleep(10)
                return True
                
        except Exception as e:
            print(f"검색 결과 처리 오류: {str(e)}")
            return False
            
    except Exception as e:
        print(f"검색 및 다운로드 오류: {str(e)}")
        return False

def main():
    """메인 함수"""
    max_retries = 3
    for attempt in range(max_retries):
        driver = None
        try:
            print(f"\n작업 시도 {attempt + 1}/{max_retries}")
            automate_opentext_workflow()
            break
            
        except Exception as e:
            print(f"시도 {attempt + 1} 실패: {str(e)}")
            if attempt < max_retries - 1:
                print("잠시 후 다시 시도합니다...")
                time.sleep(5)

if __name__ == "__main__":
    main()