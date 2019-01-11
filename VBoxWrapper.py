import virtualbox
import subprocess
import os
from sys import platform
import re
import time
import threading
import errno


# Auto-login may be enabled, but, note that this is a potential security risk as a malicious application
# running on the guest could request this information using the proper interface.
# https://www.virtualbox.org/manual/ch09.html#autologon
# Restoring from a snapshot while running is not permitted by VirtualBox.

class VBoxWrapper:
    def __init__(self):
        self.vbox = virtualbox.VirtualBox()
        self.path = self.find_vboxmanage()
        self.list_vm = self.getAvailableVMs()
        self.currentVMID = ""
        self.username = "username"
        self.password = "password"

    def find_vboxmanage(self):
        # return main path to be used in commands
        def is_64bit():
            if 'PROCESSOR_ARCHITEW6432' in os.environ:
                return True
            return os.environ['PROCESSOR_ARCHITECTURE'].endswith('64')

        selectedPath = ""
        if self.detectOSType() == "windows":
            if is_64bit():
                selectedPath = "C:\\Program Files\\Oracle\\VirtualBox\\VBoxManage.exe"
            else:
                selectedPath = "C:\\Program Files (x86)\\Oracle\\VirtualBox\\VBoxManage.exe"
        else:
            selectedPath = "VBoxManage"
        return selectedPath

    def detectOSType(self):
        # Return type of OS
        if platform == "linux" or platform == "linux2":
            return 'linux'
        elif platform == "darwin":
            return 'macos'
        elif platform == "win32":
            return 'windows'

    def start_vm(self, name):
        try:
            # Start a vm given name
            # output = self.run_command([self.path, "startvm", name])
            process = subprocess.Popen([self.path, "startvm", name], stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                       bufsize=1)
            output, _ = process.communicate()
            parts = output.split("\n")
            for part in parts:
                if part.__contains__("not find a registered machine named"):
                    print part
                    return -1
                elif part.__contains__("already locked by a session"):
                    print "Checking VM: Already running - " + self.list_vm[name]
                    return -2
                elif part.__contains__("error"):
                    print part
                    return -1
                else:
                    print part
            return 0
        except Exception, e:
            print e
            return -1

    def stop_vm(self, name):
        # Stopping vm given name
        try:
            print "Stopping VM..."
            # output = self.run_command([self.path, "controlvm", name, "poweroff"])
            process = subprocess.Popen([self.path, "controlvm", name, "poweroff"], stdout=subprocess.PIPE,
                                       stderr=subprocess.STDOUT,
                                       bufsize=1)
            output, _ = process.communicate()
            flag = 0
            parts = output.split("\n")
            for part in parts:
                if (part.__contains__("Details:")) or (part.__contains__("Context:")) or part.__contains__(
                        "is not currently running"):
                    flag = -1
                    break
            return flag
        except Exception, e:
            print e
            return -1

    def getAvailableVMs(self):
        # return a list of available vms' names
        list = {}
        result = subprocess.Popen([self.path, "list", "vms"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        output, _ = result.communicate()
        parts = output.split("\n")
        for part in parts:
            if part:
                s1 = re.search('"(.*)"', part)
                s2 = re.search('{(.*)}', part)
                list[s2.group(1)] = s1.group(1)
        return list

    def screenShotAndMoveToHost(self, vmid, hostPath):
        # Takes a screenshot and copies it to given path in host.
        print "Screenshotting to: " + hostPath
        process = subprocess.Popen([self.path, "controlvm", vmid, 'screenshotpng', hostPath], stdout=subprocess.PIPE,
                                   stderr=subprocess.STDOUT)
        output, error = process.communicate()
        flag = True
        if output:
            print output
            if output.__contains__("error:"):
                flag = False
        if error:
            print error
            if error.__contains__("error:"):
                flag = False
        # code = process.returncode
        return flag

    def timeout(self, p):
        # helper function for execute_in_vm
        # kills process
        if p.poll() is None:
            try:
                p.kill()
                # print 'Error: process taking too long to complete: please check login credentials if set yet.'
            except OSError as e:
                if e.errno != errno.ESRCH:
                    print "Process couldn't terminate."

    def execute_in_vm(self, name, command):
        # executes a program in vm given path of executable and name of vm
        process = subprocess.Popen(
            [self.path, 'guestcontrol', name, 'run', '--exe', command, '--username', self.username,
             '--password', self.password], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        t = threading.Timer(5.0, self.timeout, [process])
        t.start()
        t.join()
        # print process
        output, error = process.communicate()
        flag = True
        if output:
            print output
            if output.__contains__("error:"):
                flag = False
        if error:
            print error
            if error.__contains__("error:"):
                flag = False
        # code = process.returncode
        return flag

    def copyFromGuestToHost(self, vmid, pathInGuest, pathInHost):
        # copies a file from Guest To host, given vm ID
        process = subprocess.Popen(
            [self.path, 'guestcontrol', vmid, '--verbose', '--username', self.username, '--password', self.password,
             'copyfrom',
             '--target-directory', pathInHost, pathInGuest], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output, error = process.communicate()
        flag = True
        if output:
            print output
            if output.__contains__("error:"):
                flag = False
        if error:
            print error
            if error.__contains__("error:"):
                flag = False
        # code = process.returncode
        return flag

    def copyFromHostToGuest(self, vmid, pathInHost, pathInGuest):
        # copies a file from host to guest, given vm ID
        process = subprocess.Popen(
            [self.path, 'guestcontrol', vmid, 'copyto', '--verbose', '--username', self.username, '--password',
             self.password, pathInHost
                , '--target-directory', pathInGuest], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output, error = process.communicate()
        flag = True
        if output:
            print output
            if output.__contains__("error:"):
                flag = False
        if error:
            print error
            if error.__contains__("error:"):
                flag = False
        # code = process.returncode
        return flag

    def getNameFile(self, dir):
        # gets only file name of a long directory path
        parts = dir.split("\\")
        return parts[len(parts) - 1]

    def getLatestSnapShot(self, vmid):
        # returns latest snapshot ID
        process = subprocess.Popen([self.path, 'showvminfo', vmid], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output, error = process.communicate()
        snapID = ""
        if output.__contains__("Snapshots:"):
            for match in re.finditer(r"UUID: (.*?)\)", output):
                snapID = match.group(1)
            return snapID
        return "-1"

    def restoreSnapShot(self, vmid):
        # restores a snapshot given vm ID
        snapShotUUID = self.getLatestSnapShot(vmid)
        if snapShotUUID != "-1":
            print 'Restoring Latest Snapshot:'
            process = subprocess.Popen([self.path, 'snapshot', vmid, 'restore', snapShotUUID], stdout=subprocess.PIPE,
                                       stderr=subprocess.PIPE)
            output, error = process.communicate()
            if output:
                print output
            if error:
                print error
            if error.__contains__("error:") or output.__contains__("error:"):
                return -1
            vbw.currentVMID = vmid
            return 0
        else:
            print "No snapshots found " + snapShotUUID
        return -1

    def executeLatestSnapshot(self):
        # Executes latest snapshot
        vbw.stop_vm(vbw.currentVMID)
        vbw.restoreSnapShot(vbw.currentVMID)
        vbw.start_vm(vbw.currentVMID)

    def pollAnalysis(self, vmid, pathInGuest, pathInHost):
        # getting TextFile from guest to host to parse it.
        if os.path.exists(pathInHost):
            os.remove(pathInHost)
        if vbw.copyFromGuestToHost(vmid, pathInGuest, pathInHost):
            with open(pathInHost) as myfile:
                print (list(myfile)[-1])


# Tutorial
if __name__ == "__main__":
    vbw = VBoxWrapper()

    execFileGuest1 = "C:\\Documents and Settings\\username\\Desktop\\msnsusii.exe"
    execFileHost1 = "D:\\" + vbw.getNameFile(execFileGuest1)
    execFileHost2 = "C:\\Users\\MyPC\\Desktop\\ccsetup536.exe"
    dirGuest1 = "C:\\Documents and Settings\\username\\Desktop\\"
    pngFileHost1 = "C:\\pic1.png"
    txtFileGuest1 = "C:\\Documents and Settings\\username\\Desktop\\newtext1.txt"
    txtFileHost1 = "C:\\Users\\MyPC\\Desktop\\text1.txt"

    print "Available VMs"
    print vbw.list_vm.keys()
    if vbw.list_vm:
        vbw.currentVMID = vbw.list_vm.iterkeys().next()
        print "Starting VM: " + vbw.list_vm[vbw.currentVMID]
        retCode = vbw.start_vm(vbw.currentVMID)
        if retCode == 0 or retCode == -2:
            if retCode == 0:
                print 'Tip: Don\'t forget to enter the login credentials.'
                print 'Sleeping for 25 secs'
                time.sleep(25)
            # in this function you need to declare filename+extension of file in Host side.
            boolVal2 = vbw.copyFromGuestToHost(vbw.currentVMID, execFileGuest1, execFileHost1)
            print boolVal2
            boolVal2 = vbw.copyFromHostToGuest(vbw.currentVMID, execFileHost2, dirGuest1)
            print boolVal2
            # you need to declare filename+extension (png) in Host side.
            vbw.screenShotAndMoveToHost(vbw.currentVMID, pngFileHost1)
            # The same for copyFromGuestToHost applies here.
            vbw.pollAnalysis(vbw.currentVMID, txtFileGuest1, txtFileHost1)
            print "Sleeping for 10 secs before restoring a snapshot.."
            time.sleep(10)
            vbw.executeLatestSnapshot()
