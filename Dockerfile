FROM node:20-slim

# LibreOffice headless + fonts (Liberation = Arial-compatible, ttf-mscorefonts = Arial officiel)
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      libreoffice-impress \
      libreoffice-core \
      libreoffice-common \
      fonts-liberation \
      fonts-dejavu-core \
      fonts-noto-core \
      fontconfig \
      ca-certificates \
      curl \
      && \
    # Microsoft core fonts (Arial etc.) — accepte l'EULA
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections && \
    apt-get install -y --no-install-recommends ttf-mscorefonts-installer && \
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
