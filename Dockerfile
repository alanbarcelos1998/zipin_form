# Usar uma imagem oficial do Node.js como base
FROM node:lts

# Criar e definir o diretório de trabalho dentro do contêiner
WORKDIR /app

# Copiar o package.json e package-lock.json (se disponível)
COPY package*.json ./

# Instalar as dependências do projeto
RUN npm install

# Copiar o restante do código do aplicativo para o contêiner
COPY . .

# Expor a porta que o aplicativo vai usar
EXPOSE 3000

# Definir o comando para iniciar o aplicativo
CMD ["npm", "start"]
