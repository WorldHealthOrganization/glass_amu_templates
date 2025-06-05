# Copyright (c) 2025 World Health Organization
# SPDX-License-Identifier: BSD-3-Clause

import configparser
import logging
import os
import os.path
import sys
from zipfile import ZipFile



logging.basicConfig(filename=f"{os.path.join(os.path.curdir, "..", "postbuild.log")}", level=logging.INFO)
log = logging.getLogger(__name__)

def load_config() -> configparser.ConfigParser:
    log.info("load config")
    cfg_path = os.path.join(os.path.curdir, "..\\..", "NAMU_Template", "settings.ini")
    log.debug(cfg_path)
    cfg = configparser.ConfigParser()
    cfg.read(cfg_path)
    log.debug(cfg.sections())
    return cfg

def rename_msi_file(orig_path: str, dev_env:str, cfg: configparser.ConfigParser)->str:
    log.info("Rename msi file")
    path,ext = os.path.splitext(orig_path)
    log.debug(path)
    log.debug(ext)
    version = cfg["Version"]["Version"]
    subversion = cfg["Version"]["Subversion"]
    new_path = f"{path}-{version}v{subversion}-release{ext}" if dev_env=="Release" else f"NAMU_addin-{version}v{subversion}-debug{ext}"
    os.rename(orig_path, new_path)
    log.debug(new_path)
    return new_path

def create_zip_file(msi_file: str, dev_env:str, cfg: configparser.ConfigParser):
    log.info("Build zip file")
    version = cfg["Version"]["Version"]
    subversion = cfg["Version"]["Subversion"]
    zip_path = f"NAMU_installer-{version}v{subversion}-release.zip" if dev_env=="Release" else f"NAMU_installer-{version}v{subversion}-debug.zip"
    with ZipFile(zip_path, 'x') as fzip:
        fzip.write("setup.exe")
        fzip.write(msi_file)
    log.debug(zip_path)
    return zip_path

if __name__ == "__main__":
    try:
        args = sys.argv
        log.info(args)
        orig_path = args[1]
        msi_path = os.path.basename(orig_path)
        log.info(orig_path)
        dev_env = "Release"
        if len(args)>2:
            dev_env = args[2]
        log.info(dev_env)
        cfg = load_config()
        # new_msi_path = rename_msi_file(orig_path, dev_env, cfg)
        # log.info(f"Msi file renamed to {new_msi_path}.")
        zip_path = create_zip_file(msi_path, dev_env, cfg)
        log.info(f"Zip file {zip_path} created.")
        sys.stdout.writelines(f"Postbuild script successful=>{zip_path}.")
        sys.exit(0)
    except Exception as e:
        sys.stderr.writelines("Unexpected error")
        sys.stderr.writelines(str(e))
        sys.exit(1)


