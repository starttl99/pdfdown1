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
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(20)
    
    return driver, download_dir

def wait_for_angular_load(driver, wait):
    """Angular 앱 로드 대기"""
    try:
        # 기본 페이지 로드 대기
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
        time.sleep(3)  # 초기 대기

        # Angular 앱 로드 확인
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-root")))
        time.sleep(2)

        # ng-version attribute 확인
        wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "app-root[ng-version]"))
        time.sleep(2)

        # Angular의 로딩 인디케이터가 사라질 때까지 대기 (있는 경우)
        try:
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "loading-indicator")))
        except:
            pass  # 로딩 인디케이터가 없을 수 있음

        # 추가 안정화 대기
        time.sleep(5)
        return True
    except Exception as e:
        print(f"Angular 로드 오류: {str(e)}")
        return False

def click_with_multiple_methods(driver, element, wait):
    """여러 방법으로 클릭 시도"""
    methods = [
        lambda: driver.execute_script("""
            arguments[0].dispatchEvent(new MouseEvent('click', {
                'view': window,
                'bubbles': true,
                'cancelable': true
            }));
        """, element),
        lambda: driver.execute_script("arguments[0].click();", element),
        lambda: element.click(),
        lambda: ActionChains(driver).move_to_element(element).click().perform()
    ]
    
    for method in methods:
        try:
            method()
            time.sleep(2)
            return True
        except Exception:
            continue
    return False

def find_and_click_element(driver, wait, selectors, scroll=True):
    """요소를 찾아서 클릭하는 함수"""
    for selector in selectors:
        try:
            print(f"선택자 시도 중: {selector}")
            element = wait.until(EC.presence_of_element_located(selector))
            wait.until(EC.visibility_of_element_located(selector))
            wait.until(EC.element_to_be_clickable(selector))
            
            if scroll:
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
                time.sleep(2)
            
            if click_with_multiple_methods(driver, element, wait):
                print(f"요소 클릭 성공: {selector}")
                time.sleep(2)
                return True
            
        except Exception as e:
            print(f"요소 클릭 오류 ({selector}): {str(e)}")
            continue
    return False

def login_to_site(driver, wait):
    """로그인 수행"""
    try:
        driver.get("https://tradinggrid.opentext.com/tgoportal/#/login")
        if not wait_for_angular_load(driver, wait):
            return False

        time.sleep(3)

        # 로그인 폼 요소 찾기
        username_selectors = [
            (By.CSS_SELECTOR, "input[formcontrolname='userId']"),
            (By.XPATH, "//input[@placeholder='User ID']")
        ]
        
        password_selectors = [
            (By.CSS_SELECTOR, "input[formcontrolname='password']"),
            (By.XPATH, "//input[@placeholder='Password']")
        ]

        # 사용자 이름 입력
        for selector in username_selectors:
            try:
                username_input = wait.until(EC.presence_of_element_located(selector))
                username_input.clear()
                username_input.send_keys("youngjang@hyoseong.co.kr")
                driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", username_input)
                print("사용자 이름 입력 성공")
                break
            except Exception:
                continue

        time.sleep(1)

        # 비밀번호 입력
        for selector in password_selectors:
            try:
                password_input = wait.until(EC.presence_of_element_located(selector))
                password_input.clear()
                password_input.send_keys("Hyoseong02!")
                driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", password_input)
                print("비밀번호 입력 성공")
                break
            except Exception:
                continue

        time.sleep(1)

        # 로그인 버튼 클릭
        sign_in_selectors = [
            (By.ID, "submit"),
            (By.XPATH, "//button[contains(text(),'Sign In')]")
        ]

        if find_and_click_element(driver, wait, sign_in_selectors):
            print("로그인 버튼 클릭 성공")
            time.sleep(5)
            return True
        else:
            print("로그인 버튼 클릭 실패")
            return False

    except Exception as e:
        print(f"로그인 오류: {str(e)}")
        return False

def navigate_to_inbox(driver, wait):
    """Document Manager로 이동 후 Inbox 클릭"""
    try:
        # 홈페이지로 이동
        print("홈페이지로 이동 중...")
        driver.get("https://tradinggrid.opentext.com/tgong/#/homepage")
        time.sleep(10)  # 초기 로딩을 위한 충분한 대기 시간
        
        if not wait_for_angular_load(driver, wait):
            print("홈페이지 로드 실패")
            return False
            
        # Document Manager 버튼이 클릭 가능할 때까지 대기
        print("Document Manager 버튼 찾는 중...")
        max_attempts = 3
        document_manager_selectors = [
            (By.CSS_SELECTOR, "div.portal-tile-title.portal-tile-header-text-align"),
            (By.XPATH, "//div[contains(@class, 'portal-tile-title')]//span[contains(text(), 'Document Manager')]"),
            (By.XPATH, "//span[contains(text(), 'Document Manager')]")
        ]
        
        # Document Manager 클릭 시도
        clicked = False
        for attempt in range(max_attempts):
            try:
                if find_and_click_element(driver, wait, document_manager_selectors):
                    clicked = True
                    print("Document Manager 클릭 성공")
                    break
                else:
                    print(f"Document Manager 클릭 실패 (시도 {attempt + 1})")
                    time.sleep(5)
            except Exception as e:
                print(f"Document Manager 클릭 오류 (시도 {attempt + 1}): {str(e)}")
                if attempt < max_attempts - 1:
                    time.sleep(5)

        if not clicked:
            print("Document Manager 클릭 실패")
            return False

        # Inbox 클릭을 위한 충분한 대기 시간
        time.sleep(10)
        
        # Inbox 선택자
        inbox_selectors = [
            (By.XPATH, "//div[contains(@class, 'tile-header')]//span[text()='Inbox']"),
            (By.XPATH, "//div[contains(@class, 'portal-tile-header')]//span[text()='Inbox']"),
            (By.CSS_SELECTOR, "div.portal-tile-header span"),
            (By.XPATH, "//span[text()='Inbox']")
        ]
        
        # Inbox 클릭 재시도 로직
        print("Inbox 클릭 시도 중...")
        for attempt in range(max_attempts):
            try:
                # 페이지가 완전히 로드될 때까지 대기
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "portal-tile-wrapper")))
                
                if find_and_click_element(driver, wait, inbox_selectors, scroll=True):
                    print(f"Inbox 클릭 성공 (시도 {attempt + 1})")
                    time.sleep(10)  # 클릭 후 충분한 대기 시간
                    
                    # 성공 확인
                    try:
                        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Type']")))
                        print("Inbox 페이지 로드 확인됨")
                        return True
                    except:
                        print("Inbox 페이지 로드 실패, 재시도...")
                        continue
                else:
                    print(f"Inbox 클릭 실패 (시도 {attempt + 1})")
                    time.sleep(5)
                    
            except Exception as e:
                print(f"Inbox 클릭 오류 (시도 {attempt + 1}): {str(e)}")
                if attempt < max_attempts - 1:
                    time.sleep(5)
                else:
                    return False

        print("최대 시도 횟수 초과")
        return False
            
    except Exception as e:
        print(f"Inbox 네비게이션 오류: {str(e)}")
        return False

def search_and_download(driver, wait):
    """검색 및 PDF 다운로드 수행"""
    try:
        type_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Type']")))
        type_input.clear()
        type_input.send_keys("DELJIT")
        time.sleep(1)

        doc_id_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Document ID']")))
        doc_id_input.clear()
        doc_id_input.send_keys("5100001476")
        time.sleep(1)

        # 검색 버튼 클릭
        search_selectors = [
            (By.XPATH, "//button[contains(text(),'Search')]"),
            (By.CSS_SELECTOR, "button.search-button")
        ]
        if not find_and_click_element(driver, wait, search_selectors):
            print("검색 버튼 클릭 실패")
            return False

        time.sleep(3)
        
        # 검색 결과 확인 및 PDF 생성
        try:
            edi_order_row = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[contains(., '5100001476')]")))
            generate_pdf_button = edi_order_row.find_element(By.XPATH, ".//button[contains(text(),'GENERATE PDF')]")
            
            if click_with_multiple_methods(driver, generate_pdf_button, wait):
                print("PDF 생성 버튼 클릭 성공")
                time.sleep(10)
                return True
            else:
                print("PDF 생성 버튼 클릭 실패")
                return False
                
        except Exception as e:
            print(f"검색 결과 처리 오류: {str(e)}")
            return False
            
    except Exception as e:
        print(f"검색 및 다운로드 오류: {str(e)}")
        return False

def automate_opentext_workflow():
    """전체 작업 흐름 실행"""
    driver, download_dir = create_driver()
    wait = WebDriverWait(driver, 45)

    try:
        print("로그인 시도 중...")
        if not login_to_site(driver, wait):
            print("❌ 로그인 실패")
            return

        print("Inbox로 이동 중...")
        if not navigate_to_inbox(driver, wait):
            print("❌ Inbox 이동 실패")
            return

        print("문서 검색 및 다운로드 중...")
        if not search_and_download(driver, wait):
            print("❌ 검색 및 다운로드 실패")
            return

        print("✅ 작업 완료")
        
    except Exception as e:
        print(f"❌ 작업 중 오류 발생: {str(e)}")
    finally:
        driver.quit()

def main():
    """메인 함수"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            print(f"\n작업 시도 {attempt + 1}/{max_retries}")
            automate_opentext_workflow()
        except Exception as e:
            print(f"시도 {attempt + 1} 실패: {str(e)}")
            if attempt < max_retries - 1:
                print("잠시 후 다시 시도합니다...")
                time.sleep(5)
            else:
                print("최대 시도 횟수를 초과했습니다.")