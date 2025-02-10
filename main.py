import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
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
                        print(f"스크롤 중 오류 발생 (무시): {str(e)}")
                
                # 여러 클릭 방법 시도
                try:
                    driver.execute_script("arguments[0].click();", element)
                except Exception:
                    try:
                        element.click()
                    except Exception:
                        try:
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

def navigate_to_inbox(driver, wait):
    """
    Document Manager 페이지에서 Inbox를 Full XPath 방식을 사용하여 클릭하는 함수.
    Document Manager 클릭은 기존과 동일하게 진행되며, 그 후 Inbox 요소에 대해
    실제 사용자가 클릭할 때 발생하는 이벤트(마우스 다운 → 마우스 업 → 클릭)를 순차적으로 재현합니다.
    """
    try:
        print("Document Manager 진입을 위한 클릭 시도 중...")
        # Document Manager 클릭 (기존 방식)
        doc_manager_selectors = [
            (By.XPATH, "/html/body/div/homepage/homepage-content/div/ot-tile-container/div/div[1]/div/div[2]/div[3]/ot-tile/div/ot-tile-content/div/div[1]/div")
        ]
        print("Document Manager 선택자 시도 중:", doc_manager_selectors[0])
        doc_manager = wait.until(EC.element_to_be_clickable(doc_manager_selectors[0]))
        
        if doc_manager:
            driver.execute_script("arguments[0].scrollIntoView(true);", doc_manager)
            time.sleep(2)
            driver.execute_script("arguments[0].click();", doc_manager)
            print("Document Manager 클릭 성공")
        else:
            print("Document Manager 요소를 찾을 수 없습니다")
            return False

        print("페이지 로드 대기 중...")
        time.sleep(30)

        # Inbox Full XPath를 사용한 클릭 (마우스 이벤트 순차 재현)
        print("Inbox Full XPath 사용하여 클릭 시도 중...")
        inbox_xpath = "/html/body/app-root/dm-root/dm-layout/div[1]/dm-banner/document-summary/div/div/div[1]/dm-summary-list/ul/li"
        inbox = wait.until(EC.presence_of_element_located((By.XPATH, inbox_xpath)))
        driver.execute_script("arguments[0].scrollIntoView(true);", inbox)
        time.sleep(2)

        driver.execute_script("""
            var target = arguments[0];
            var rect = target.getBoundingClientRect();
            var x = rect.left + rect.width / 2;
            var y = rect.top + rect.height / 2;
            
            var mousedown = new MouseEvent('mousedown', {
                bubbles: true,
                cancelable: true,
                view: window,
                clientX: x,
                clientY: y
            });
            target.dispatchEvent(mousedown);
            
            var mouseup = new MouseEvent('mouseup', {
                bubbles: true,
                cancelable: true,
                view: window,
                clientX: x,
                clientY: y
            });
            target.dispatchEvent(mouseup);
            
            var click = new MouseEvent('click', {
                bubbles: true,
                cancelable: true,
                view: window,
                clientX: x,
                clientY: y
            });
            target.dispatchEvent(click);
        """, inbox)

        print("Inbox 클릭 이벤트(마우스 다운/업/클릭) 시도 완료")
        time.sleep(5)
        return True

    except Exception as e:
        print("전체 네비게이션 프로세스 오류:", e)
        return False

def search_and_download(driver, wait):
    """검색 및 PDF 다운로드 수행"""
    try:
        # 1. Filter list 버튼 클릭
        filter_list_xpath = "/html/body/app-root/dm-root/dm-layout/div[1]/dm-banner/dm-listing-table/div/div[2]/div/nav/a"
        filter_button = wait.until(EC.element_to_be_clickable((By.XPATH, filter_list_xpath)))
        driver.execute_script("arguments[0].click();", filter_button)
        time.sleep(2)

        # 2. Document ID 입력
        doc_id_xpath = "/html/body/app-root/dm-root/dm-layout/div[1]/dm-banner/dm-listing-table/div/div[2]/div[1]/dm-listing-search/ul/div[1]/perfect-scrollbar/div/div[1]/div[3]/input"
        doc_id_input = wait.until(EC.presence_of_element_located((By.XPATH, doc_id_xpath)))
        doc_id_input.clear()
        doc_id_input.send_keys("5100001476")
        time.sleep(2)

        # 3. Type Deljit 클릭
        type_deljit_xpath = "/html/body/app-root/dm-root/dm-layout/div[1]/dm-banner/dm-listing-table/div/div[2]/div[1]/dm-listing-search/ul/div[1]/perfect-scrollbar/div/div[1]/div[8]/ul/li[4]/div/div"
        type_deljit = wait.until(EC.element_to_be_clickable((By.XPATH, type_deljit_xpath)))
        driver.execute_script("arguments[0].click();", type_deljit)
        time.sleep(2)

        # 4. Apply 버튼 클릭
        apply_xpath = "/html/body/app-root/dm-root/dm-layout/div[1]/dm-banner/dm-listing-table/div/div[2]/div[1]/dm-listing-search/ul/div[2]/div[2]/button[1]"
        apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, apply_xpath)))
        driver.execute_script("arguments[0].click();", apply_button)
        time.sleep(5)

        # 검색 결과 처리
        try:
            row = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[contains(., '5100001476')]")))
            if find_and_click_element(driver, wait, [(By.XPATH, ".//button[contains(text(),'GENERATE PDF')]")]):
                print("PDF 생성 버튼 클릭 성공")
                time.sleep(10)
                return True
        except Exception as e:
            print("검색 결과 처리 오류:", e)
            return False
            
    except Exception as e:
        print("검색 및 다운로드 오류:", e)
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

        # 로그인 후 Document Manager 페이지로 이동 후, Inbox 접근 (Full XPath + 마우스 이벤트 재현)
        if navigate_to_inbox(driver, wait):
            print("Document Manager 진입 및 Inbox 접근 성공")
            # 검색 및 다운로드 실행
            if search_and_download(driver, wait):
                print("검색 및 다운로드 성공")
            else:
                print("검색 및 다운로드 실패")
        else:
            print("Document Manager 진입 및 Inbox 접근 실패")

    finally:
        if driver:
            driver.quit()

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
            print(f"시도 {attempt + 1} 실패:", e)
            if attempt < max_retries - 1:
                print("잠시 후 다시 시도합니다...")
                time.sleep(5)

if __name__ == "__main__":
    main()
