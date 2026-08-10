FROM node:24-alpine

RUN apk add --no-cache font-dejavu font-liberation font-noto-cjk

WORKDIR /app
COPY dist ./dist
COPY dist-server ./dist-server

ENV GENOFFICE_WASM_DIR=/app/dist-server/wasm

EXPOSE 80

CMD ["node", "dist-server/server.mjs"]
