#![windows_subsystem = "windows"]

use anyhow::{bail, Context, Result};
use native_windows_derive as nwd;
use native_windows_gui as nwg;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use winreg::{enums::*, RegKey};

use nwd::NwgUi;
use nwg::NativeUi;

fn create_link<P: AsRef<Path>>(
    target: P,
    args: &str,
    title: &str,
    output: P,
    cwd: Option<P>,
    icon: Option<P>,
) -> Result<()> {
    let file = std::env::temp_dir().join("tmp.vbs");
    std::fs::write(
        &file,
        format!(
            "
Set objWS = WScript.CreateObject(\"WScript.Shell\")
strLinkFile = \"{}\"
Set objLink = objWS.CreateShortcut(strLinkFile)
objLink.TargetPath = \"{}\"
objLink.Arguments = \"{}\"
objLink.Description = \"{}\"
{}
{}
objLink.Save
    ",
            output.as_ref().to_str().unwrap(),
            target.as_ref().to_str().unwrap().replace(r"\\?\", ""),
            args,
            title,
            if let Some(i) = icon {
                format!(
                    "objLink.IconLocation = \"{}\"",
                    i.as_ref().to_str().unwrap()
                )
            } else {
                "".to_owned()
            },
            if let Some(dir) = cwd {
                format!(
                    "objLink.WorkingDirectory = \"{}\"",
                    dir.as_ref().to_str().unwrap()
                )
            } else {
                "".to_owned()
            }
        ),
    )?;
    let output = std::process::Command::new("cscript")
        .args(&["/nologo", file.to_str().unwrap()])
        .creation_flags(0x00000008)
        .output()?;
    if !output.stderr.is_empty() {
        let msg = String::from_utf8_lossy(&output.stderr);
        bail!("{}", msg);
    }
    std::fs::remove_file(&file)?;
    Ok(())
}

#[derive(Default, NwgUi)]
pub struct BcmlInstaller {
    #[nwg_control(title: "BCML Installer", flags: "WINDOW|VISIBLE", size: (250, 226))]
    #[nwg_events( OnWindowClose: [BcmlInstaller::close])]
    window: nwg::Window,

    #[nwg_control(text: "Thank you for downloading BCML. Set your\r\ninstall options below:", size: (344, 32), position: (7, 7))]
    thank_you: nwg::Label,

    #[nwg_control(text: "Create Start Menu shortcut", size: (344, 16), position: (7, 46), check_state: nwg::CheckBoxState::Checked)]
    start_bcml: nwg::CheckBox,

    #[nwg_control(text: "Create desktop shortcut", size: (344, 16), position: (7, 76))]
    desktop_bcml: nwg::CheckBox,

    #[nwg_control(text: "Create updater shortcut in Start Menu", size: (344, 16), position: (7, 106), check_state: nwg::CheckBoxState::Checked)]
    update_bcml: nwg::CheckBox,

    #[nwg_control(text: "Create uninstall shortcut in Start Menu", size: (344, 16), position: (7, 136), check_state: nwg::CheckBoxState::Checked)]
    uninstall_bcml: nwg::CheckBox,

    #[nwg_control(text: "Add BCML Python to PATH", size: (344, 16), position: (7, 166))]
    do_path: nwg::CheckBox,

    #[nwg_control(text: "Close", position: (7, 196), size: (96, 24))]
    #[nwg_events( OnButtonClick: [BcmlInstaller::close] )]
    close_button: nwg::Button,
    #[nwg_control(text: "OK", position: (110, 196), size: (96, 24))]
    #[nwg_events( OnButtonClick: [BcmlInstaller::do_it] )]
    ok_button: nwg::Button,
}

impl BcmlInstaller {
    fn do_it(&self) {
        let runner = || -> Result<()> {
            let cwd = std::env::current_dir()?;

            for file in glob::glob("**/*.dll")
                .unwrap()
                .chain(glob::glob("**/*.exe").unwrap())
                .flat_map(|f| f.ok())
            {
                let ads = format!(
                    "{}:Zone.Identifier",
                    file.file_name().unwrap().to_str().unwrap()
                );
                std::fs::remove_file(file.with_file_name(&ads)).ok();
            }

            let pypath = cwd.join("pythonw.exe").canonicalize()?;
            let start_dir = PathBuf::from(std::env::var("APPDATA")?)
                .join(r"Microsoft\Windows\Start Menu\Programs\BCML");
            let icon = cwd.join(r"Lib\site-packages\bcml\data\bcml.ico");
            std::fs::create_dir_all(&start_dir)?;

            if let nwg::CheckBoxState::Checked = self.start_bcml.check_state() {
                create_link(
                    &pypath,
                    "-m bcml",
                    "BCML",
                    &start_dir.join("BCML.lnk"),
                    Some(&cwd),
                    Some(&icon),
                )?;
            }
            if let nwg::CheckBoxState::Checked = self.desktop_bcml.check_state() {
                create_link(
                    &pypath,
                    "-m bcml",
                    "BCML",
                    &PathBuf::from(std::env::var("USERPROFILE")?).join("Desktop\\BCML.lnk"),
                    Some(&cwd),
                    Some(&icon),
                )?;
            }
            if let nwg::CheckBoxState::Checked = self.update_bcml.check_state() {
                create_link(
                    &pypath,
                    "-m pip install -U bcml",
                    "Update BCML",
                    &start_dir.join("Update.lnk"),
                    Some(&cwd),
                    None,
                )?;
            }
            if let nwg::CheckBoxState::Checked = self.uninstall_bcml.check_state() {
                create_link(
                    std::env::current_exe()?,
                    "--uninstall",
                    "Uninstall BCML",
                    start_dir.join("Uninstall.lnk"),
                    Some(cwd.clone()),
                    None,
                )?;
            }
            if let nwg::CheckBoxState::Checked = self.do_path.check_state() {
                let hkcu = RegKey::predef(HKEY_CURRENT_USER);
                let (env, _) = hkcu.create_subkey("Environment")?;
                let env_path: String = env.get_value("PATH")?;
                env.set_value("PATH", &format!("{};{}", &env_path, cwd.to_str().unwrap()))?;
            };
            Ok(())
        };
        if let Err(e) = runner() {
            self.error(&format!("{}", e));
        } else {
            if let nwg::MessageChoice::Ok = nwg::message(&nwg::MessageParams {
                title: "BCML Installed",
                content: "BCML was successfully installed. Do you want to exit?",
                buttons: nwg::MessageButtons::OkCancel,
                icons: nwg::MessageIcons::Question,
            }) {
                self.close();
            }
        }
    }
    fn error(&self, content: &str) {
        nwg::modal_error_message(&self.window, "Error", content);
    }
    fn close(&self) {
        nwg::stop_thread_dispatch();
    }
}

fn main() {
    if std::env::args()
        .find(|a| a.contains("--uninstall"))
        .is_none()
    {
        nwg::init().expect("Failed to init Native Windows GUI");

        let mut font = nwg::Font::default();
        nwg::Font::builder()
            .size(16)
            .family("Segoe UI")
            .build(&mut font)
            .expect("Failed to build font");
        nwg::Font::set_global_default(Some(font));

        let _app = BcmlInstaller::build_ui(Default::default()).expect("Failed to build UI");
        nwg::dispatch_thread_events();
    } else {
        nwg::init().expect("Failed to init Native Windows GUI");

        if let nwg::MessageChoice::Ok = nwg::message(&nwg::MessageParams {
            title: "Uninstall BCML",
            content: "Are you sure you want to uninstall BCML?",
            buttons: nwg::MessageButtons::OkCancel,
            icons: nwg::MessageIcons::Question,
        }) {
            let runner = || -> Result<()> {
                let start_dir = PathBuf::from(std::env::var("APPDATA")?)
                    .join(r"Microsoft\Windows\Start Menu\Programs\BCML");
                let desktop_link =
                    PathBuf::from(std::env::var("USERPROFILE")?).join(r"Desktop\BCML.lnk");

                std::fs::remove_file(&desktop_link).ok();
                std::fs::remove_dir_all(&start_dir).ok();

                for item in std::fs::read_dir(std::env::current_exe()?.parent().unwrap())?
                    .filter_map(|i| i.ok().map(|i| i.path()))
                {
                    if item.is_dir() {
                        std::fs::remove_dir_all(item).ok();
                    } else if item.is_file() {
                        std::fs::remove_file(&item).ok();
                    }
                }

                Ok(())
            };
            if let Err(e) = runner() {
                nwg::error_message(
                    "Uninstall Error",
                    &format!("There was an error uninstalling BCML: {}", e),
                );
            } else {
                nwg::simple_message("BCML Uninstalled", "Successfully uninstalled BCML!");
            }
        }
        nwg::stop_thread_dispatch();
    }
}
