@echo off
@secedit /analyze /db %1_cfg.sec /cfg xpsp2_r3.inf /log %1_sec.log
find "�� ᮮ⢥�����" %1_sec.log >%1_sec!.log

