@echo off
REM Define o titulo da janela do console
title Atualizador de Repositorio GitHub

echo =================================================================
echo.
echo           ATUALIZANDO ARQUIVOS PARA O GITHUB...
echo.
echo =================================================================
echo.

REM Passo 1: Adiciona todos os arquivos novos e modificados.
echo --- Adicionando todos os arquivos (git add .)...
git add .
echo.

REM Passo 2: Cria um "commit" com uma mensagem padrao (incluindo data e hora).
echo --- Criando commit (git commit)...
git commit -m "Atualizacao de arquivos - %date% %time%"
echo.

REM Passo 3: Envia as alteracoes para o repositorio remoto no GitHub.
echo --- Enviando para o GitHub (git push)...
git push
echo.

echo =================================================================
echo.
echo      PROCESSO FINALIZADO! VERIFIQUE SEU GITHUB.
echo.
echo =================================================================
echo.
pause