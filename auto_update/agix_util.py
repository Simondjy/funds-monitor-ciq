import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def setup_driver():
    """设置Chrome浏览器驱动"""
    try:
        logger.info("正在设置Chrome浏览器驱动...")
        
        # Chrome选项配置 - 只使用无头模式
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        
        # 设置用户代理
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36")
        
        # 其他必要的选项
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # 自动下载并设置ChromeDriver
        service = Service(ChromeDriverManager().install())
        
        # 创建WebDriver实例
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # 执行脚本来隐藏webdriver属性
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        logger.info("Chrome浏览器驱动设置成功")
        return driver
        
    except Exception as e:
        logger.error(f"设置Chrome浏览器驱动失败: {e}")
        return None

def get_agix_shares_outstanding():
    """获取AGIX ETF的shares outstanding数据"""
    driver = None
    try:
        driver = setup_driver()
        if not driver:
            return None
        
        url = "https://kraneshares.com/agix/"
        logger.info(f"正在获取AGIX ETF数据从: {url}")
        
        # 访问网页
        driver.get(url)
        
        # 等待页面加载
        time.sleep(3)
        
        # 等待数据表格加载
        wait = WebDriverWait(driver, 20)
        
        try:
            # 等待数据表格出现
            data_table = wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "data_table"))
            )
            logger.info("数据表格加载成功")
            
        except TimeoutException:
            logger.error("等待数据表格超时")
            return None
        
        # 查找Shares Outstanding行
        rows = data_table.find_elements(By.TAG_NAME, "tr")
        
        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) == 2:
                    key = cells[0].text.strip()
                    if key == "Shares Outstanding":
                        value = cells[1].text.strip()
                        logger.info(f"成功获取Shares Outstanding: {value}")
                        return value
            except Exception as e:
                continue
        
        logger.error("未找到Shares Outstanding数据")
        return None
        
    except Exception as e:
        logger.error(f"获取数据时出错: {e}")
        return None
    
    finally:
        # 关闭浏览器
        if driver:
            driver.quit()
            logger.info("浏览器已关闭")

def main():
    """主函数 - 测试工具函数"""
    print("开始测试AGIX爬虫工具...")
    
    shares_outstanding = get_agix_shares_outstanding()
    
    if shares_outstanding:
        print("✓ 数据获取成功!")
        print(f"Shares Outstanding: {shares_outstanding}")
    else:
        print("✗ 数据获取失败!")

if __name__ == "__main__":
    main()
