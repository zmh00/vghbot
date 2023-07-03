import json
import urllib.request
import subprocess
import sys

class Updater_github():
    """
    flet為多線程程序，流程為初始化確認是否最新+設定參數 => 啟動flet介面 => 非同步啟動下載 => 下載完執行+關閉flet
    """
    def __init__(self, owner:str, repo:str, target_name:str, version_tag:str) -> None:
        self.owner = owner # github owner
        self.repo = repo # github repository
        self.target_name = target_name # search for specific filename
        self.version_tag= version_tag # version_tag in this program
        
        # The following items are retrieved after the connection
        self.version_tag_latest = None
        self.target_fullname = None
        self.download_url = None
        self.filename = None
        # self.progressbar = None


    def start(self):
        if self.is_latest() == False:
            self.download()
            return False # 給呼叫更新的原程式一個訊息
        else:
            return True # 給呼叫更新的原程式一個訊息


    def get_info(self):
        """
        Get version_tag_latest, target_fullname, download_url
        """
        latest_url = f"https://api.github.com/repos/{self.owner}/{self.repo}/releases/latest"
        with urllib.request.urlopen(latest_url) as response:
            res = json.loads(response.read().decode())

            if 'tag_name' in res:
                self.version_tag_latest = res['tag_name']
                latest_assets = res['assets']
            else:
                print("Failed to retrieve latest version")
                sys.exit('Failed to retrieve latest version')

        for asset in latest_assets:
            if self.target_name.upper() in asset['name'].upper(): # insensitive name match
                self.download_url = asset['browser_download_url']
                self.target_fullname = asset['name']
                break
        
    
    def is_latest(self):
        if self.version_tag_latest == None:
            self.get_info()
        if self.version_tag < self.version_tag_latest: # check whether the local program is the latest with string comparison in python
            print("Not the latest version!")
            return False
        else:
            print("Already the latest version!") 
            return True


    def download(self):     
        with urllib.request.urlopen(self.download_url) as response:
            total_size = int(response.info().get('Content-Length', 0))
            downloaded_size = 0
            filename, extension = self.target_fullname.split('.') # split the filename 
            filename = f'{filename}({self.version_tag_latest}).{extension}' # reset the file name and add version tag
            self.filename = filename
            with open(filename, 'wb') as f:
                while True:
                    buffer = response.read(8192) # chunk size
                    if not buffer:
                        break
                    
                    # Write the chunk to the file
                    f.write(buffer)

                    # Update the downloaded size and display progress
                    downloaded_size += len(buffer)
                    print(f"Progress: {downloaded_size/total_size*100 :.2}%", end='\r')
        
        subprocess.Popen(self.filename) # 執行新程式
        