FROM golang:1.24.3-alpine AS builder

WORKDIR /src

COPY go.mod go.sum ./
RUN go mod download

COPY . .
RUN CGO_ENABLED=0 GOOS=linux go build -o /out/berrio .

FROM alpine:3.22

WORKDIR /app

RUN apk add --no-cache ca-certificates tzdata

COPY --from=builder /out/berrio /app/berrio
COPY public /app/public
COPY docs /app/docs

EXPOSE 8081

CMD ["/app/berrio"]
