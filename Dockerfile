# Built on the NAS GitHub runner, pushed to git.filipkin.com/filip/argotrack,
# pulled by Coolify on the cloud box. See .github/workflows/deploy.yml.

FROM node:22-slim AS build
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY tsconfig.json ./
COPY src ./src
RUN npm run build   # tsc -> dist/

FROM node:22-slim
WORKDIR /app
# curl + wget: Coolify's health check execs one of them inside the container.
RUN apt-get update && apt-get install -y --no-install-recommends curl wget \
    && rm -rf /var/lib/apt/lists/*
ENV NODE_ENV=production
COPY package.json package-lock.json ./
RUN npm ci --omit=dev
COPY --from=build /app/dist ./dist
COPY public ./public
EXPOSE 3000
HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=5 \
    CMD curl -fsS http://localhost:3000/health || exit 1
CMD ["node", "dist"]
