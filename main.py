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

        # 로그인 후 Document Manager로 이동
        if navigate_to_inbox(driver, wait):
            print("Document Manager 진입 성공")
            # 검색 및 다운로드 실행
            if search_and_download(driver, wait):
                print("검색 및 다운로드 성공")
            else:
                print("검색 및 다운로드 실패")
        else:
            print("Document Manager 진입 실패")

    finally:
        if driver:
            driver.quit()

def navigate_to_inbox(driver, wait):
    try:
        print("Document Manager로 이동 시도 중...")
        
        # Document Manager 클릭 (기존 코드 유지)
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
            
        print("\n페이지 로드 대기 중...")
        time.sleep(30)
        
        print("\n=== DOM 분석 시작 ===")
        
        # 1. 모든 li 요소 검색 및 상세 정보 출력
        print("\n[ li 요소 검색 중... ]")
        all_li_elements = driver.find_elements(By.TAG_NAME, "li")
        print(f"총 {len(all_li_elements)}개의 li 요소 발견\n")
        
        for idx, li in enumerate(all_li_elements, 1):
            try:
                print(f"\n=== Li Element #{idx} ===")
                print("▶ 기본 속성:")
                print(f"  - ID: {li.get_attribute('id')}")
                print(f"  - Class: {li.get_attribute('class')}")
                print(f"  - Role: {li.get_attribute('role')}")
                print(f"  - Text Content: {li.text}")
                
                print("▶ HTML 구조:")
                print(f"  {li.get_attribute('outerHTML')}")
                
                print("▶ 위치 정보:")
                location = li.location
                size = li.size
                print(f"  - Position: x={location['x']}, y={location['y']}")
                print(f"  - Size: width={size['width']}, height={size['height']}")
                
                print("▶ 상태 정보:")
                print(f"  - Displayed: {li.is_displayed()}")
                print(f"  - Enabled: {li.is_enabled()}")
                
                # XPath 정보
                print("▶ XPath:")
                xpath = driver.execute_script("""
                    function getPathTo(element) {
                        if (element.id !== '')
                            return "//*[@id='" + element.id + "']";
                        if (element === document.body)
                            return element.tagName.toLowerCase();
                        var ix = 0;
                        var siblings = element.parentNode.childNodes;
                        for (var i = 0; i < siblings.length; i++) {
                            var sibling = siblings[i];
                            if (sibling === element)
                                return getPathTo(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';
                            if (sibling.nodeType === 1 && sibling.tagName === element.tagName)
                                ix++;
                        }
                    }
                    return getPathTo(arguments[0]);
                """, li)
                print(f"  - Generated XPath: {xpath}")
                
                print("-" * 50)
                
            except Exception as e:
                print(f"Element #{idx} 분석 중 오류: {str(e)}")
                continue
        
        print("\n=== DOM 분석 완료 ===")
        
        # 2. Inbox 요소 찾기 시도
        print("\nInbox 요소 찾기 시도 중...")
        target_element = None
        
        for li in all_li_elements:
            try:
                if (li.get_attribute('id') == 'dm_dibHeader' or 
                    'inbox' in li.text.lower() or 
                    'dm_dib' in (li.get_attribute('class') or '')):
                    target_element = li
                    print("잠재적인 Inbox 요소 발견:")
                    print(f"- ID: {li.get_attribute('id')}")
                    print(f"- Class: {li.get_attribute('class')}")
                    print(f"- Text: {li.text}")
                    break
            except:
                continue
        
        if target_element:
            print("\nInbox 요소 클릭 시도...")
            driver.execute_script("arguments[0].scrollIntoView(true);", target_element)
            time.sleep(2)
            driver.execute_script("arguments[0].click();", target_element)
            print("Inbox 클릭 성공")
            return True
        else:
            print("\nInbox 요소를 찾을 수 없습니다.")
            return False

    except Exception as e:
        print(f"\n전체 네비게이션 프로세스 오류: {str(e)}")
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
            # 검색 결과 행에서 PDF 생성 버튼 찾기
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