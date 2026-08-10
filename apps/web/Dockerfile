FROM node:24-alpine

WORKDIR /app
COPY dist ./dist
COPY dist-server ./dist-server

EXPOSE 80

CMD ["node", "dist-server/server.mjs"]
