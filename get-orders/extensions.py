from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.action_chains import ActionChains
import sys

driver = None
action = None

def set_driver(driver_value,action_value=None):
    global driver,action
    driver = driver_value
    action = action_value

# Custom Made By seth
def find(self,method,selector,forceList=False):
    content = self
    elements = []
    if method == 'class':
        elements = content.find_elements_by_class_name(selector)
    elif method == 'text':
        elements = content.find_elements_by_link_text(selector)
    elif method == 'text*':
        elements = content.find_elements_by_partial_link_text(selector)
    elif method == 'text+':
        elements = content.find_elements_by_xpath('//*[text()[contains(., "'+selector+'")]]')
    elif method == 'text~':
        elements = content.find_elements_by_xpath('.//*[text()[contains(., "'+selector+'")]]')
    elif method == 'tag':
        elements = content.find_elements_by_tag_name(selector)
    elif method == 'name':
        elements = content.find_elements_by_name(selector)
    elif method == 'id':
        elements = content.find_elements_by_id(selector)
    elif method == 'xpath':
        elements = content.find_elements_by_xpath(selector)
    else:
        elements = content.find_elements_by_xpath(".//*[@"+method+"='"+selector+"']")

    if len(elements) == 0:
        if forceList:
            return elements
        else:
            raise Exception("No element matching criteria: %s = '%s'" % (method,selector)) 
            return None
    elif len(elements) == 1:
        if forceList:
            return elements
        else:
            return elements[0]
    else:
        return elements

def strong_click(self,method,selector):
    try:
        if method == 'text':
            xpath = ".//*[text()='"+selector+"']"
        if method == 'xpath':
            xpath = selector
        else:
            xpath = ".//*[@"+method+"='"+selector+"']"
        if self.__class__.__name__ == 'WebDriver':
            driver.execute_script("document.evaluate(\""+xpath+"\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()")
        else:    
            driver.execute_script("document.evaluate(\""+xpath+"\", arguments[0], null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()",self)
    except Exception as e:
        print(e)
        raise Exception('Error while strong_click "%s" = "%s"' % (method,selector))

def flag(self,color='lightgreen'):
    driver.execute_script("arguments[0].style += 'background: "+color+";'", self) 
    return self

def send(self,text):
    self.clear()
    self.send_keys(text)

def delete(self):
    driver.execute_script('arguments[0].parentNode.removeChild(arguments[0]);', self)

def up(self,num=1):
    elem = self
    for i in range(num):
        elem = elem.find_element_by_xpath('..')
    return elem
    
exts = ['find','flag','send','up','delete','strong_click']

for ext in exts:
    exec('WebDriver.'+ext+' = '+ext)
    exec('WebElement.'+ext+' = '+ext)

# -------- Action Chains -------

def do(self):
    self.perform()
    action.reset_actions()

ActionChains.do = do


    
