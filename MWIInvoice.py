import aiohttp, asyncio,os, queue, time,json
from openpyxl import load_workbook, Workbook
from bs4 import BeautifulSoup as bs
from concurrent.futures import ThreadPoolExecutor

USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) ' \
             'Chrome/98.0.4758.102 Safari/537.36 Edg/98.0.1108.62'
LIMIT = 3
TIMEOUT = 600

class MWIInvoice:
    def __init__(self,gui_queue,startdate,enddate):
        self.gui_queue = gui_queue
        self.client = self.username = self.password = self.accountid = None
        self.startdate = startdate
        self.enddate = enddate

    async def auth_login(self):
        url = 'https://store.mwiah.com/api/mwi/userwrite/signin'

        headers = {
            'Connection': 'keep-alive',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Microsoft Edge";v="98"',
            'Accept': 'application/json, text/plain, */*',
            'DNT': '1',
            'Content-Type': 'application/json;charset=UTF-8',
            'sec-ch-ua-mobile': '?0',
            'User-Agent': USER_AGENT,
            'sec-ch-ua-platform': '"Windows"',
            'Origin': 'https://store.mwiah.com',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Dest': 'empty',
            'Referer': 'https://store.mwiah.com/sign-in?page=%2F',
            'Accept-Language': 'en-US,en;q=0.9',
        }

        # data = '{"userName":"jkshah@analytix.com","userPassword":"u#zTt9Zu"}'
        data = '{"userName":"--username--","userPassword":"--password--"}'
        data = data.replace("--username--", self.username).replace("--password--", self.password)
        async with self.sema:
            async with self.session.post(url,headers=headers,data=data) as request:
                response = await request.content.read()
                content = response.decode('utf-8')
                json_content = json.loads(content)
                if json_content.get("isValid") and not json_content.get("hasInvalidCredentials"):
                    return True
                else:
                    return False

    async def login(self):
        url = 'https://store.mwiah.com/'

        headers = {
            'Connection': 'keep-alive',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Microsoft Edge";v="98"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'Upgrade-Insecure-Requests': '1',
            'DNT': '1',
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-User': '?1',
            'Sec-Fetch-Dest': 'document',
            'Referer': 'https://store.mwiah.com/sign-in?page=%2F',
            'Accept-Language': 'en-US,en;q=0.9',
        }
        async with self.sema:
            async with self.session.get(url,headers=headers) as request:
                response = await request.content.read()
                content = response.decode('utf-8')
                login_content = bs(content,"html.parser")
                title = login_content.find('title').text
                if title.strip() == "MWI Animal Health | Dashboard":
                    return True
                else:
                    return False

    async def change_account(self):
        url = 'https://store.mwiah.com/user/change-account'
        headers = {
            'Connection': 'keep-alive',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Microsoft Edge";v="98"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'Upgrade-Insecure-Requests': '1',
            'DNT': '1',
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;'
                      'q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-User': '?1',
            'Sec-Fetch-Dest': 'document',
            'Accept-Language': 'en-US,en;q=0.9',
        }

        params = (
            ('accountId', f'{self.accountid}'),
        )
        try:
            async with self.sema:
                async with self.session.get(url,headers=headers,params=params) as request:
                    response = await request.content.read()
                    content = response.decode('utf-8')
                    html_content = bs(content, "html.parser")
                    account = html_content.find("span", attrs={"data-test-id", "account-menu-id"})
                    accountid = account.text
                    if accountid != self.accountid:
                        self.gui_queue.put({'Status': f'{self.client} Client not found'})
                        return False
                    else:
                        return True

        except Exception as e:
            print(e)

    async def get_invoice_page(self):
        url = 'https://marketplace.vgpvet.com/order-history/invoices'
        headers = {
            'Connection': 'keep-alive',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Microsoft Edge";v="98"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'Upgrade-Insecure-Requests': '1',
            'DNT': '1',
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;'
                      'q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-User': '?1',
            'Sec-Fetch-Dest': 'document',
            'Referer': 'https://marketplace.vgpvet.com/dashboard',
            'Accept-Language': 'en-US,en;q=0.9',
        }
        async with self.sema:
            async with self.session.get(url, headers=headers) as request:
                response = await request.content.read()
                content = response.decode('utf-8')
                html_content = bs(content,"html.parser")
                title = html_content.find('title').text
                if title.strip() == "Veterinary Growth Partners | Invoices":
                    return True
                else:
                    return False

    async def get_invoices(self):
        url = 'https://marketplace.vgpvet.com/api/mwi/orders/invoicesearch'

        headers = {
            'Connection': 'keep-alive',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Microsoft Edge";v="98"',
            'Accept': 'application/json, text/plain, */*',
            'DNT': '1',
            'Content-Type': 'application/json;charset=UTF-8',
            'sec-ch-ua-mobile': '?0',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36 Edg/98.0.1108.62',
            'sec-ch-ua-platform': '"Windows"',
            'Origin': 'https://marketplace.vgpvet.com',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Dest': 'empty',
            'Referer': 'https://marketplace.vgpvet.com/order-history/invoices?size=50&sortKey=orderhistory-invoices%3A%3Amostrecent&start=2022-01-02&end=2022-03-04&page=1',
            'Accept-Language': 'en-US,en;q=0.9',
        }

        data = {"pageNumber":1,"pageSize":50,"sortValue":"orderhistory-invoices::mostrecent","startDate":self.startdate,"endDate":self.enddate,"facets":{},"term":""}
        async with self.sema:
            async with self.session.post(url,headers=headers,data=data) as request:
                response = await request.content.read()
                content = response.decode('utf8')



    async def download_process(self):
        timeout = aiohttp.ClientTimeout(total=TIMEOUT)
        conn = aiohttp.TCPConnector(limit=5, limit_per_host=5)
        self.sema = asyncio.Semaphore(LIMIT)
        async with aiohttp.ClientSession(connector=conn, timeout=timeout) as self.session:
            auth_login = await self.auth_login()
            if not auth_login:
                self.gui_queue.put({'Status':'Unable to authenticate'}) if self.gui_queue else None

            login = await self.login()
            if not login:
                self.gui_queue.put({'Status':'Unable to login'}) if self.gui_queue else None

            change = await self.change_account()
            if not change:
                self.gui_queue.put({'Status': 'Unable to change account'}) if self.gui_queue else None

            invoice_page = await self.get_invoice_page()
            if not invoice_page:
                self.gui_queue.put({'Status': 'Unable to open invoice page'}) if self.gui_queue else None

            invoice = await self.get_invoices()
            if not invoice:
                self.gui_queue.put({'Status': 'Unable to fetch invoices'}) if self.gui_queue else None


    def start_download(self):
        try:
            loop = asyncio.new_event_loop()
            future = asyncio.ensure_future(self.download_process(),loop=loop)
            loop.run_until_complete(future)
            return future.result()
        except Exception as e:
            pass

class RunMWI:
    def __init__(self):
        self.gui_queue = queue.Queue()
    def run(self):
        run_start = time.perf_counter()
        setting = 'MWISettingSheet.xlsx'
        setting_wb = load_workbook(setting, data_only=True, read_only=True)
        setting_ws = setting_wb['Creds'].values
        setting_data = [list(row) for row in setting_ws if row]
        mwi = MWIInvoice(self.gui_queue,"2022-01-02","2022-03-04")
        for row in setting_data:
            if len(row) >= 4:
                if row[1] == "no account id found":
                    self.gui_queue.put({'Status':f'Account ID not Found for client {row[0]}'})if self.gui_queue else None
                    continue
                else:
                    mwi.client = row[0]
                    mwi.accountid = row[1]
                    mwi.username = row[2]
                    mwi.password = row[3]
                    mwi.start_download()

        run_end = time.perf_counter()
        time_taken = time.strftime("%H:%M:%S", time.gmtime(int(run_end - run_start)))
        print(f'Time Taken = {time_taken}')
        self.gui_queue.put({"status": f"Time Taken {time_taken}"})







if __name__ == '__main__':
    run = RunMWI()
    run.run()