FROM node:latest

# Create app directory
WORKDIR /usr/src/app

# Install app dependencies
# A wildcard is used to ensure both package.json AND package-lock.json are copied
# where available (npm@5+)
COPY package*.json ./



# Bundle app source
COPY . .

RUN apt update -y
RUN apt upgrade -y
RUN apt install npm -y
RUN npm install
RUN npx playwright install-deps
RUN npx playwright install


EXPOSE 7125
CMD [ "node", "index.js" ]