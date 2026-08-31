# Gerador de Tópico 1 — EBD em Foco

Aplicação web que recebe o `.pptx` de uma lição completa e gera
automaticamente o "Tópico 1" (slide de amostra), em `.pptx` e `.pdf`,
seguindo exatamente o padrão que você já usa manualmente.

## Como funciona (regra de corte)

1. Tudo **antes** do primeiro slide de seção ("I. ...") é mantido: capa,
   texto áureo/principal, resumo, objetivos, texto bíblico, introdução.
2. O **Tópico I inteiro** é mantido (bumper + todo o conteúdo).
3. Dos demais tópicos (II, III, IV...), mantém-se **só o slide de
   título** — o conteúdo é descartado.
4. Conclusão, revisão e o encerramento original são descartados.
5. Um **pacote fixo de slides promocionais** é anexado ao final —
   `assets/promo_jovens.pptx` ou `assets/promo_adultos.pptx`, conforme
   o tipo escolhido. Esses dois arquivos foram extraídos dos seus
   próprios exemplos (`Topico-1-Licao_de_jovens.pptx` /
   `..._adultos.pptx`) e podem ser abertos e editados no PowerPoint a
   qualquer momento — o app sempre usa a versão mais recente desses
   dois arquivos.
6. Na capa, a caixa de texto "Nome do Professor" é substituída por
   "SLIDE DE AMOSTRA / APENAS TÓPICO I".

Essa lógica foi validada gerando o Tópico 1 a partir das duas lições
completas que você enviou (Sansão/jovens e Paulo/adultos) e comparando
slide a slide com os arquivos que você já tinha pronto — o resultado
bateu 100%.

⚠️ **Pré-requisito do arquivo de entrada:** os slides de seção
("I.", "II.", "III."...) precisam ser slides só com o título (texto +
imagem de fundo, no máximo 3 elementos), igual ao padrão que você já
usa. Se algum dia o modelo do slide mudar bastante, me avise para eu
ajustar a regra de detecção.

## Rodando localmente (Windows, para testar antes de publicar)

```bash
pip install -r requirements.txt
python app.py
```

Abra `http://localhost:5000` no navegador. Para a conversão em PDF
funcionar localmente você precisa ter o **LibreOffice** instalado e o
comando `soffice` disponível no PATH (no Windows, o instalador do
LibreOffice já adiciona isso automaticamente na maioria dos casos; se
não funcionar, adicione a pasta `program` da instalação, ex.
`C:\Program Files\LibreOffice\program`, à variável PATH do Windows).

## Publicando online (recomendado: Render.com, tem plano gratuito)

O jeito mais simples é publicar como um **Web Service com Docker** —
o `Dockerfile` já cuida de instalar o LibreOffice no servidor.

1. Crie um repositório no GitHub e suba esta pasta inteira (`git init`,
   `git add .`, `git commit -m "primeira versão"`, `git push`).
2. Crie uma conta em https://render.com (dá pra logar com o GitHub).
3. Clique em **New +** → **Web Service**.
4. Selecione o repositório que você acabou de subir.
5. Em "Environment", escolha **Docker** (o Render detecta o
   `Dockerfile` automaticamente).
6. Plano: **Free** serve para começar (o primeiro acesso após um
   tempo sem uso pode demorar ~30s para "acordar" o servidor — normal
   em planos gratuitos).
7. Clique em **Create Web Service** e aguarde o build (a primeira vez
   demora alguns minutos porque instala o LibreOffice).
8. Pronto — o Render te dá uma URL tipo
   `https://seu-app.onrender.com`, que já é a aplicação publicada.

Alternativas equivalentes (mesmo Dockerfile funciona): **Railway.app**
e **Fly.io** — ambos também têm suporte nativo a Docker e planos de
entrada gratuitos/baratos.

### Se preferir sem Docker (não recomendado)

Praticamente todo host "sem Docker" (PythonAnywhere, Vercel, etc.) não
permite instalar o LibreOffice, que é obrigatório para gerar o PDF.
Por isso a rota com Docker é a mais simples e confiável para este
projeto.

## Estrutura do projeto

```
app.py                 # aplicação Flask (rotas de upload/geração/download)
topico1_builder.py      # motor: detecta os bumpers e monta o Tópico 1
slide_copy.py            # utilitário para copiar slides entre apresentações
templates/index.html     # página de upload
templates/resultado.html # página de download do resultado
assets/promo_jovens.pptx   # pacote fixo de slides promocionais (jovens)
assets/promo_adultos.pptx  # pacote fixo de slides promocionais (adultos)
requirements.txt
Dockerfile
```

## Manutenção

- **Trocar os slides promocionais:** edite/substitua
  `assets/promo_jovens.pptx` ou `assets/promo_adultos.pptx` no
  PowerPoint normalmente e suba de novo (git push) — não precisa mexer
  em código.
- **Arquivos temporários:** cada geração cria uma pasta com um ID
  único em `tmp_uploads/` e `tmp_outputs/`; o app limpa
  automaticamente pastas com mais de 2 horas a cada nova geração, para
  não acumular disco no servidor.
- **Limite de tamanho de upload:** 150 MB (ajustável em `app.py`,
  constante `MAX_CONTENT_LENGTH`).
