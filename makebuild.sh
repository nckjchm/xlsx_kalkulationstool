#!/bin/bash


pyinstaller ./kalkulationsanwendung.py \
    --noconfirm --log-level=WARN \
    --clean \
    --onefile \
    --add-data="resourcen/Kalkulationsvorlage.xlsx:resourcen/" \
    --add-data="resourcen/logo.png:resourcen/" \
    -w

    