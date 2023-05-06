from logging import exception

from selenium import webdriver
import time
from selenium.webdriver.common.by import By
import openpyxl
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import TimeoutException


#Find element based on input xpath and return true if element exists
def check_exists_by_xpath(xpath):
    if len(driver.find_elements(By.XPATH, xpath)) > 0:
        return True
    return False
driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://test-warranty.cambria-local.co.uk/Identity/Account/Login")
time.sleep(2)
# from workbook:
wk = openpyxl.load_workbook("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
sheet1 = wk.active
user_cols = sheet1['F4':'F15']
col = 4
b_column = 4

for c_column in range(4,15):
    user = sheet1['f' + str(c_column)]
    print("user is:", user.value)
    Email = driver.find_element(By.ID, "Input_Email")
    Email.clear()
    Password = driver.find_element(By.ID, "Input_Password")
    Password.clear()
    password_col = sheet1['g' + str(b_column)]
    b_column += 1
    print("user name is : ", user.value)
    print("password is :", password_col.value)
    if user.value is None: #aasheesh@fcile.com
        if password_col.value is None:
            driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[5]/button").click()
            message1 = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[2]/span/span").text
            message2 = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[3]/span/span").text
            print("Message1 is :", message1)
            print("Message2 is:", message2)
            if "The Email field is required." in message1 and "The Password field is required." in message2:
                print("Invalid credentials")
                sheet1['j' + str(col)].value = "pass"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
        else:
            Password.send_keys(password_col.value)
            driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[5]/button").click()
            error_message = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[1]/ul/li").text
            print("Message is :", error_message)
            if "Email field is required" in error_message:
                print("Email field is not entered")
                sheet1['j' + str(col)].value = "pass"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
    else:
        if password_col.value is None:
            Email.send_keys(user.value)
            #Password.clear()
            driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[5]/button").click()
            error_message = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[3]/span/span").text
            print("Error message is:", error_message)
            if "The Password field is required" in error_message:
                print("Password is not entered")
                sheet1['j' + str(col)].value = "pass"
                col+= 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
        else:
            Email.send_keys(user.value)
            Password.send_keys(password_col.value)
            #Email.clear()
            #Password.clear()
            #Email.send_keys(user.value)
            #Password.send_keys(password_col.value)
            #print("User able to edit credentials")
            #sheet1['j' + str(col)].value = "pass"
            #col += 1
            #wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            # Try to login
            driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[5]/button").click()
            if check_exists_by_xpath("/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[1]/ul/li"):
                error_message = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[1]/ul/li").text
                print("Error message is:",error_message)
                if "The Email field is not a valid e-mail address." in error_message:
                    print("Email address is invalid")
                    sheet1['j' + str(col)].value = "pass"
                    col += 1
                    wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")

            elif check_exists_by_xpath("/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[1]/ul/li"):
                if "Invalid login attempt." in error_message:
                    print("Invalid credential")
                    sheet1['j' + str(col)].value = "pass"
                    col += 1
                    wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")

            elif driver.title:
                print("Title is:", driver.title)
                print("Login Success")
                sheet1['j' + str(col)].value = "pass"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")

                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                if check_exists_by_xpath("/html/body/div/div[2]/div/div/div[1]/p/a"):
                    print("Logged out successfully")
                    sheet1['j' + str(col)].value = "pass"
                    col += 1
                    wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                    driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a").click()
                else:
                    print("user is not logged out")

user_cols = sheet1['F15':'F25']
for c_column in range(15, 25):
    user = sheet1['f' + str(c_column)]
    print("user is:", user.value)
    Email = driver.find_element(By.ID, "Input_Email")
    Email.clear()
    Password = driver.find_element(By.ID, "Input_Password")
    Password.clear()
    password_col = sheet1['g' + str(b_column)]
    b_column += 1
    print("user name is : ", user.value)
    print("password is :", password_col.value)
    Email.send_keys(user.value)
    Password.send_keys(password_col.value)
    time.sleep(2)

    if c_column == 15:
        Email.clear()
        Email.send_keys("asheesh@facileconsulting.com")
        time.sleep(1)
        Password.clear()
        Password.send_keys("Asheesh@12345")
        time.sleep(1)
        print("User able to edit credentials")
        sheet1['j' + str(col)].value = "pass"
        col += 1
        wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
        continue

    driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[5]/button").click()
    settings = driver.find_element(By.XPATH,"/html/body/div/div[2]/header/div/div[3]/a/i")
    settings.click()
#uservalue = username.get_attribute()
#usernamevalue = username.get_attribute("value")
    if c_column == 16:
        username = driver.find_element(By.ID, "Username")
        phonenum = driver.find_element(By.ID, "Input_PhoneNumber")
        submit = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[3]/button")
        profile = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        pwd = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        if username and phonenum and submit and profile and pwd:
            print("profile page is success")
            sheet1['j' + str(col)].value = "pass"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
            continue
        else:
            print("Profile page is incorrect")
            sheet1['j' + str(col)].value = "fail"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
            continue
    if c_column == 17:
        username = driver.find_element(By.ID, "Username").get_attribute("value")
        print("username is :", username)
        print("User value is :", user.value)

        if username == user.value.strip():
            print("username is same as logged in user")
            sheet1['j' + str(col)].value = "pass"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
            print("logout success")
            continue
        else:
            print("user name is not same as logged in user")
            sheet1['j' + str(col)].value = "fail"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
            print("logout success")
            continue

    if c_column == 18:
        phonenum = driver.find_element(By.ID, "Input_PhoneNumber")
        phonenum.send_keys("9878755")
        sheet1['j' + str(col)].value = "pass"
        col += 1
        wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
        logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
        logout.click()
        login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
        login_button.click()
        continue
    if c_column == 19:
        submit = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[3]/button")
        submit.click()
        if check_exists_by_xpath("/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[4]/div"):
            profile_update = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[4]/div").text
            print("Profile updated message :", profile_update)
            sheet1['j' + str(col)].value = "pass"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
        else:
            print("some issue occurred")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
    if c_column == 20:
        change_password = driver.find_element(By.ID,"change-password")
        change_password.click()
        current_password = driver.find_element(By.ID, "Input_OldPassword")
        new_password = driver.find_element(By.ID, "Input_NewPassword")
        confirm_new_password = driver.find_element(By.ID, "Input_ConfirmPassword")
        update_password = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/button")
        profile = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        password = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        print("Title is:",driver.title)
        if current_password and new_password and confirm_new_password and update_password and profile and password:
            if driver.title:
                print("Test case is passed")
                sheet1['j' + str(col)].value = "pass"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
                login_button.click()
                continue
            else:
                print("Test case is failed")
                sheet1['j' + str(col)].value = "fail"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
                login_button.click()
                continue
        else:
            print("Test case is failed")
            sheet1['j' + str(col)].value = "fail"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
            continue
    if c_column == 21:
        change_password = driver.find_element(By.ID, "change-password")
        change_password.click()
        current_password = driver.find_element(By.ID, "Input_OldPassword")
        new_password = driver.find_element(By.ID, "Input_NewPassword")
        confirm_new_password = driver.find_element(By.ID, "Input_ConfirmPassword")
        update_password = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/button")
        profile = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        password = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        current_password.send_keys("Test@123")
        new_password.send_keys("Asheesh@1234")
        confirm_new_password.send_keys("Asheesh@1234")
        update_password.click()
        if check_exists_by_xpath("/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[1]/ul/li"):
            error_message= driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[1]/ul/li").text
            if "Incorrect password." in error_message:
                print("error message is:",error_message)
                print("Test case is passed")
                sheet1['j' + str(col)].value = "pass"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
                login_button.click()
                continue
            else:
                print("Test case is failed")
                sheet1['j' + str(col)].value = "fail"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
                login_button.click()
                continue
        else:
            print("Tet case failed")
            sheet1['j' + str(col)].value = "fail"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
            continue

    if c_column == 22:
        change_password = driver.find_element(By.ID, "change-password")
        change_password.click()
        current_password = driver.find_element(By.ID, "Input_OldPassword")
        new_password = driver.find_element(By.ID, "Input_NewPassword")
        confirm_new_password = driver.find_element(By.ID, "Input_ConfirmPassword")
        update_password = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/button")
        profile = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        password = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        current_password.send_keys(user.value)
        new_password.send_keys("Asheesh@1234")
        confirm_new_password.send_keys("Asheesh@12")
        update_password.click()
        if check_exists_by_xpath("/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[4]/span/span"):
            error_message = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[4]/span/span").text
            print("error message is:",error_message)
            if "The new password and confirmation password do not match." in error_message:
                print("Test case is passed")
                sheet1['j' + str(col)].value = "pass"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
                login_button.click()
                continue
            else:
                print("Test case is failed")
                sheet1['j' + str(col)].value = "fail"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
                login_button.click()
                continue
        else:
            print("Tet case failed. Unexpected error")
            sheet1['j' + str(col)].value = "fail"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
            continue
    if c_column == 23:
        change_password = driver.find_element(By.ID, "change-password")
        change_password.click()
        current_password = driver.find_element(By.ID, "Input_OldPassword")
        new_password = driver.find_element(By.ID, "Input_NewPassword")
        confirm_new_password = driver.find_element(By.ID, "Input_ConfirmPassword")
        update_password = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/button")
        profile = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        password = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        current_password.send_keys(user.value)
        new_password.send_keys("Test@12345")
        confirm_new_password.send_keys("Test@12345")
        np = driver.find_element(By.ID, "Input_NewPassword").get_attribute("value")
        cnp = driver.find_element(By.ID, "Input_ConfirmPassword").get_attribute("value")
        print("np is:", np)
        print("cnp is :", cnp)
        update_password.click()
        if check_exists_by_xpath("/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/div"):
            error_message = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/div").text
            print("Message is:",error_message)
            if "Your password has been changed." in error_message:
                print("Test case is passed")
                sheet1['j' + str(col)].value = "pass"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
                login_button.click()
                continue
            else:
                print("Test case is failed")
                sheet1['j' + str(col)].value = "fail"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
                login_button.click()
                continue
        else:
            print("Tet case failed. Unexpected error")
            sheet1['j' + str(col)].value = "fail"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
            continue

    if c_column == 24:
        change_password = driver.find_element(By.ID, "change-password")
        change_password.click()
        current_password = driver.find_element(By.ID, "Input_OldPassword")
        new_password = driver.find_element(By.ID, "Input_NewPassword")
        confirm_new_password = driver.find_element(By.ID, "Input_ConfirmPassword")
        update_password = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/button")
        profile = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        password = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div/div[2]/a[1]")
        current_password.send_keys(user.value)
        new_password.send_keys("Testing123")
        confirm_new_password.send_keys("Testing123")
        np = driver.find_element(By.ID, "Input_NewPassword").get_attribute("value")
        cnp = driver.find_element(By.ID, "Input_ConfirmPassword").get_attribute("value")
        print("np is:", np)
        print("cnp is :", cnp)
        update_password.click()
        time.sleep(2)
        if check_exists_by_xpath("/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[1]/ul/li"):
            message = driver.find_element(By.XPATH,'/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[1]/ul/li').text
            #message1 = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[1]/ul/li[2]").text
            #message2 = driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/form/div[1]/ul/li[3]").text
            print("message is:",message)
            if "Passwords must have at least one non alphanumeric character." in message:
                print("Test case is passed")
                sheet1['j' + str(col)].value = "pass"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
                login_button.click()
                continue
            else:
                print("Test case is failed")
                sheet1['j' + str(col)].value = "fail"
                col += 1
                wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
                logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
                logout.click()
                login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
                login_button.click()
                continue
        else:
            print("Tet case failed. Unexpected error")
            sheet1['j' + str(col)].value = "fail"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()
            continue

user_cols = sheet1['F32':'F34']
col = 32
for c_column in range(32, 34):
    user = sheet1['f' + str(c_column)]
    print("user is:", user.value)
    Email = driver.find_element(By.ID, "Input_Email")
    Email.clear()
    Password = driver.find_element(By.ID, "Input_Password")
    Password.clear()
    password_col = sheet1['g' + str(b_column)]
    b_column += 1
    print("user name is : ", user.value)
    print("password is :", password_col.value)
    Email.send_keys(user.value)
    Password.send_keys(password_col.value)
    time.sleep(2)
    driver.find_element(By.XPATH,"/html/body/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/section/form/div[5]/button").click()
    if c_column == 32:
        driver.find_element(By.XPATH,"/html/body/div/div[1]/a").click() #home
        time.sleep(2)
        print("title is:", driver.title)
        if driver.title:
            print("User able to navigate home from left panel")
        else:
            print("Home is not redirected")
        driver.find_element(By.XPATH,"/html/body/div/div[1]/nav/a[1]").click() #policy search
        time.sleep(2)
        if driver.title:
            print("User able to navigate search from left panel")
        else:
            print("search is not redirected")
        driver.find_element(By.XPATH,"/html/body/div/div[1]/nav/a[2]").click() #client
        if driver.title:
            print("User able to navigate client from left panel")
        else:
            print("client is not redirected")
        driver.find_element(By.XPATH,"/html/body/div/div[1]/nav/a[3]").click() #policy dox
        if driver.title:
            print("User able to navigate policy docs from left panel")
        else:
            print("policy docs is not redirected")
        driver.find_element(By.XPATH,"/html/body/div/div[1]/nav/a[5]").click() #new policy
        if driver.title:
            print("User able to navigate new policy from left panel")
        else:
            print("new policy is not redirected")
        driver.find_element(By.XPATH,"/html/body/div[1]/div[1]/nav/a[6]").click() #products
        if driver.title:
            print("User able to navigate products from left panel")
        else:
            print("products is not redirected")
        driver.find_element(By.XPATH,"/html/body/div[1]/div[1]/nav/a[7]").click() #suppliers
        if driver.title:
            print("User able to navigate suppliers from left panel")
        else:
            print("suppliers is not redirected")
        driver.find_element(By.XPATH,"/html/body/div/div[1]/nav/a[8]").click() #sys users
        if driver.title:
            print("User able to navigate systm users from left panel")
        else:
            print("system users is not redirected")
        driver.find_element(By.XPATH,"/html/body/div/div[1]/nav/a[4]").click()  #invoices
        print("title is:",driver.title)
        if driver.title:
            print("User able to navigate invoices from left panel")
            sheet1['j' + str(col)].value = "pass"
            col += 1
            wk.save("C:\\Users\\FC-\\Desktop\\Test Cases\\Warranty Application_TC.xlsx")
            logout = driver.find_element(By.XPATH, "/html/body/div/div[2]/header/div/div[3]/form/button/i")
            logout.click()
            login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/p/a")
            login_button.click()

        else:
            print("invoices not redirected")

    #if c_column == 33:
    #    home.send_keys(Keys.COMMAND + 't')
    #    policy_search.send_keys(Keys.COMMAND + 't')
    #    client_home.send_keys(Keys.COMMAND + 't')
    #    policy_dox.send_keys(Keys.COMMAND + 't')
    #    new_policy.send_keys(Keys.COMMAND + 't')
    #    products.send_keys(Keys.COMMAND + 't')
    #    suppliers.send_keys(Keys.COMMAND + 't')
    #    system_users.send_keys(Keys.COMMAND + 't')
    #    invoices.send_keys(Keys.COMMAND + 't')
    #    print("driver.title")
    #    if driver.title:
    #        print("User able to navigate any modules from left panel in home page")
    #    else:
    #        print("Issue :", driver.title)















