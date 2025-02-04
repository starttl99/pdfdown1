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

def check_login_success(driver, wait):
   """로그인 성공 여부를 확인하는 함수"""
   try:
       wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'logged-in')]")))
       return True
   except TimeoutException:
       return False

def create_driver():
   """WebDriver 설정 및 생성"""
   # 다운로드 폴더 설정
   download_dir = os.path.join(os.getcwd(), "downloads")
   if not os.path.exists(download_dir):
       os.makedirs(download_dir)

   # Chrome 옵션 설정
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
   
   # WebDriver 생성
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
       time.sleep(10)  # 로그인 후 충분한 대기 시간

       # 4. Homepage 이동 및 Document Manager 클릭
       print("홈페이지로 이동 중...")
       driver.get("https://tradinggrid.opentext.com/tgong/#/homepage")
       time.sleep(10)  # 페이지 로딩을 위한 충분한 대기 시간
       
       try:
           # Document Manager 클릭 시도
           document_manager_selectors = [
               (By.XPATH, "//span[contains(text(), 'Document Manager')]"),
               (By.XPATH, "//div[contains(text(), 'Document Manager')]"),
               (By.XPATH, "//*[contains(text(), 'Document Manager')]"),
               (By.PARTIAL_LINK_TEXT, "Document Manager"),
               (By.XPATH, "//a[contains(@href, 'documents')]"),
               (By.CSS_SELECTOR, "a[href*='documents']")
           ]

           print("Document Manager 찾는 중...")
           document_manager = None
           for selector in document_manager_selectors:
               try:
                   element = wait.until(EC.element_to_be_clickable(selector))
                   if element:
                       driver.execute_script("arguments[0].click();", element)
                       print("Document Manager 클릭 성공")
                       document_manager = element
                       time.sleep(5)
                       break
               except Exception as e:
                   print(f"선택자 {selector} 시도 실패: {str(e)}")
                   continue

           if not document_manager:
               print("Document Manager를 찾을 수 없어 직접 URL로 이동합니다.")
               driver.get("https://tradinggrid.opentext.com/tgong/#/documents")
               time.sleep(5)

           # 페이지 전환 확인
           wait.until(EC.url_contains("documents"))
           print("Documents 페이지로 전환 완료")
           time.sleep(5)

       except Exception as e:
           print(f"Document Manager 접근 실패: {str(e)}")
           return

       # 5. Inbox 클릭
       inbox_selectors = [(By.XPATH, "//span[text()='Inbox']")]
       inbox_element = find_and_click_element(driver, wait, inbox_selectors)
       if not inbox_element:
           print("Inbox 버튼을 찾을 수 없습니다.")
           return

       time.sleep(5)  # Inbox 클릭 후 대기

       # 6. 검색 필터 적용
       try:
           # Type 입력
           type_input = wait.until(EC.presence_of_element_located(
               (By.XPATH, "//input[@placeholder='Type']")))
           type_input.clear()
           type_input.send_keys("DELJIT")
           time.sleep(1)

           # Document ID 입력
           doc_id_input = wait.until(EC.presence_of_element_located(
               (By.XPATH, "//input[@placeholder='Document ID']")))
           doc_id_input.clear()
           doc_id_input.send_keys("5100001476")
           time.sleep(1)

           # 검색 버튼 클릭
           search_selectors = [(By.XPATH, "//button[contains(text(),'Search')]")]
           search_button = find_and_click_element(driver, wait, search_selectors)
           if not search_button:
               print("검색 버튼을 찾을 수 없습니다.")
               return

           time.sleep(3)

           # 7. 검색 결과 확인 및 PDF 생성
           edi_order_row = wait.until(EC.presence_of_element_located(
               (By.XPATH, "//tr[contains(., '5100001476')]")))
           
           generate_pdf_button = edi_order_row.find_element(
               By.XPATH, ".//button[contains(text(),'GENERATE PDF')]")
           driver.execute_script("arguments[0].click();", generate_pdf_button)
           time.sleep(10)

           print("PDF 다운로드가 완료되었습니다. 다운로드 폴더를 확인하세요:", download_dir)

       except Exception as e:
           print(f"검색 및 PDF 생성 중 오류 발생: {str(e)}")
           return

   except Exception as e:
       print(f"작업 중 오류 발생: {str(e)}")
   finally:
       try:
           if driver:
               driver.quit()
       except Exception:
           pass

def main():
   max_retries = 3
   retry_count = 0
   retry_delay = 10  # 재시도 간격을 10초로 증가
   
   while retry_count < max_retries:
       try:
           print(f"\n시도 {retry_count + 1} 시작...")
           automate_opentext_workflow()
           break
       except Exception as e:
           print(f"시도 {retry_count + 1} 실패: {str(e)}")
           retry_count += 1
           if retry_count < max_retries:
               print(f"{retry_delay}초 후 재시도합니다...")
               time.sleep(retry_delay)
           else:
               print("최대 재시도 횟수를 초과했습니다.")

if __name__ == "__main__":
   main()