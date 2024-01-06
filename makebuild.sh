#!/bin/bash


pyinstaller ./kalkulationsanwendung.py \
    --noconfirm --log-level=WARN \
    --clean \
    --onedir \
    --add-data="resourcen/Kalkulationsvorlage.xlsx:resourcen/" \
    -w

    