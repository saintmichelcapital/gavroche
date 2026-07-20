FROM node:20-slim

# LibreOffice headless + polices open source (metric-compatibles avec Arial/Times/Courier)
# NB : on n'installe PAS ttf-mscorefonts-installer (fragile car dépend d'un téléchargement
# externe qui échoue régulièrement). Liberation Sans est metric-compatible avec Arial ;
# le PPTX rendu par LibreOffice sera donc visuellement très proche du rendu PowerPoint,
# et à l'ouverture réelle dans PowerPoint la vraie Arial sera utilisée.
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      libreoffice-impress \
      libreoffice-core \
      libreoffice-common \
      fonts-liberation \
      fonts-liberation2 \
      fonts-dejavu-core \
      fonts-noto-core \
      fontconfig \
      ca-certificates \
      && \
    fc-cache -f -v && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Installe les deps Node en cache
COPY package*.json ./
RUN npm install --omit=dev

# Copie le code
COPY . .

# Le port par défaut de Render
EXPOSE 3000

# Démarrage du serveur Express
CMD ["node", "server/index.js"]
