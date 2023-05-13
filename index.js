const {globalShortcut} = require('electron');
const {AbstractTypiePackage, TypieRowItem} = require('typie-sdk');
const {execFile} = require('child_process');
const path = require('path');

const listOfExecutablePaths = [
    path.join(__dirname, './VirtualDesktopInsider.exe'),
    path.join(__dirname, './VirtualDesktop11.exe'),
    path.join(__dirname, './VirtualDesktop11-21H2.exe'),
    path.join(__dirname, './VirtualDesktop.exe'),
    path.join(__dirname, './VirtualDesktop2022.exe'),
    path.join(__dirname, './VirtualDesktop2016.exe'),
    path.join(__dirname, './VirtualDesktopServer2022.exe'),
];

class VirtualDesktops extends AbstractTypiePackage {
    constructor(win, config, pkgPath) {
        super(win, config, pkgPath);
        this.win = win;
        this.packageName = 'VirtualDesktops';
        this.selectedExecutable = "";

        this.desktopToActiveMap = {};

        globalShortcut.register('Alt+1', () => this.switchToDesktopNumber(1));
        globalShortcut.register('Alt+2', () => this.switchToDesktopNumber(2));
        globalShortcut.register('Alt+3', () => this.switchToDesktopNumber(3));
        globalShortcut.register('Alt+4', () => this.switchToDesktopNumber(4));
        globalShortcut.register('Alt+5', () => this.switchToDesktopNumber(5));
        globalShortcut.register('Alt+6', () => this.switchToDesktopNumber(6));
        globalShortcut.register('Alt+7', () => this.switchToDesktopNumber(7));
        globalShortcut.register('Alt+8', () => this.switchToDesktopNumber(8));
        globalShortcut.register('Alt+9', () => this.switchToDesktopNumber(9));

        this.bruteForceSelectExecutable(0);
    }

    bruteForceSelectExecutable(index) {
        console.debug(`${this.packageName} trying windows executable: ${listOfExecutablePaths[index]}`);
        if (index > listOfExecutablePaths.length - 1) {
            console.error(`${this.packageName} did not found any executble for virtual desktop on this windows version!`);
            return;
        }
        try {
            execFile(listOfExecutablePaths[index], ["/List"], (error, stdout, stderr) => {
                if (error) {
                    this.bruteForceSelectExecutable(index + 1);
                    return;
                }
                this.selectedExecutable = listOfExecutablePaths[index];
                console.info(`${this.packageName} found virtual desktop executable: ${this.selectedExecutable}`);
                return;
            });
        } catch (e) {
            this.bruteForceSelectExecutable(index + 1);
        }
    }

    async switchToDesktopNumber(num) {
        // before switch - save the current hwnd foreground window
        const hwnd = await this.getCurrentActiveWindow();
        const index = num - 1;

        if (hwnd) {
            const activeDesktopNumber = await this.getCurrentActiveDesktop();
            this.desktopToActiveMap[`desk_${activeDesktopNumber}`] = hwnd;
            if (index == activeDesktopNumber) {
                return;
            }
        }

        console.log("switching to index: ", index);
        this.callExecutable([`-Switch:${index}`], (stdout) => {
            console.log("switch: ", stdout);
            if (this.desktopToActiveMap[`desk_${index}`]) {
                this.setFocusToWindow(this.desktopToActiveMap[`desk_${index}`]);
            }
        });
    }

    async setFocusToWindow(hwnd) {
        const item = {
            db: 'global',
            t: 'SwitchTo',
            p: hwnd,
        }
        const rowItem = TypieRowItem.create(item);
        console.log("setting focus", rowItem.p, this.desktopToActiveMap);
        return this.typie.switchTo(rowItem).go();
    }

    async getCurrentActiveWindow() {
        return new Promise((resolve) => {
            this.windowManager([`-getCurrent`], (stdout) => {
                console.log("current active hwnd:", stdout);
                if (stdout) {
                    resolve(stdout.trim());
                } else {
                    resolve(0);
                }
            });
        });
    }

    async getCurrentActiveDesktop() {
        return new Promise((resolve, reject) => {
            this.callExecutable([`-gcd`], (stdout) => {
                const res = stdout.trim().match(/\s(\d*?)\)$/);
                if (res) {
                    console.log("current active desktop:", res[1]);
                    resolve(res[1]);
                } else {
                    reject();
                }
            });
        });
    }

    activate(pkgList, item, cb) {
        console.log(`${this.packageName} switch to: ${item.title}`);
        this.callExecutable([`-Switch:${item.title}`], (stdout) => {
            const desktops = this.getDesktops(stdout);
            if (desktops.length > 0) {
                this.populate(desktops);
            }
        });
    }

    enterPkg(pkgList, item, cb) {
        this.callExecutable(["/List"], (stdout) => {
            const desktops = this.getDesktops(stdout);
            if (desktops.length > 0) {
                this.populate(desktops);
            }
        });
    }

    callExecutable(args, callback) {
        if (!this.selectedExecutable) {
            return;
        }

        execFile(this.selectedExecutable, args, (error, stdout, stderr) => {
            callback && callback(stdout);
            // if (error) {
            //     this.handleExecutionError(error, stdout, stderr);
            // } else {
            //     callback && callback(stdout);
            // }
        });
    }

    windowManager(args, callback) {
        const windowManagerPath = path.join(__dirname, './window-manager.exe');
        execFile(windowManagerPath, args, (error, stdout, stderr) => {
            if (error) {
                this.handleExecutionError(error, stdout, stderr);
            } else {
                callback && callback(stdout);
            }
        });
    }

    handleExecutionError(error, stdout, stderr) {
        // do nothing!
        // if (this.selectedExecutable) {
        //     console.error(error, stdout, stderr);
        // }
        console.error("err:", stderr);
    }

    getDesktops(stdout) {
        const lines = stdout.split("\r\n");
        const desktops = [];
        for (let line of lines) {
            if (line.includes(" (Wallpaper: ")) {
                let splits = line.split(" (Wallpaper: ");
                let name = splits[0];
                let isVisible = false;
                if (name.includes(" (visible)")) {
                    isVisible = true;
                    name = name.split(" (visible)")[0];
                }
                desktops.push({
                    isVisible: isVisible,
                    name: name,
                    wallpaper: splits[1].slice(0, -1)
                });
            }
        }
        return desktops;
    }

    populate(desktops) {
        const itemsArray = [];

        for (let desktop of desktops) {
            itemsArray.push(
                new TypieRowItem(desktop.name)
                    .setDB(this.packageName)
                    .setPackage(this.packageName)
                    .setDescription(`switch to: ${desktop.name}`)
                    .setIcon(desktop.wallpaper));
        }

        // this.win.send('resultList', itemsArray);

        this.win.send("resultList", {data: itemsArray, length: itemsArray.length, err: 0});

        // this.typie.multipleInsert(itemsArray).go()
        //     .then(data => {
        //         console.info("VirtualDesktops plugin done adding", data);
        //         this.typie.getRows(10).orderBy('unixTime').desc().go()
        //                 .then(res => {
        //                     this.win.send('resultList', res);
        //                     this.win.show();
        //                 })
        //                 .catch(err => console.log(err));
        //     })
        //     .catch(err => console.error("VirtualDesktops plugin insert error", err));
    }
}

module.exports = VirtualDesktops;

