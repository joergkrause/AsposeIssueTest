#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/azure-functions/dotnet:4 AS base
## Aspose support (see: https://docs.aspose.com/pdf/net/docker/)
RUN apt-get update && apt-get install -y --allow-unauthenticated libgdiplus libc6-dev
#RUN sed -i'.bak' 's/$/ contrib/' /etc/apt/sources.list
# Fonts for deploy on Linux
COPY ./AsposeIssueTest/build/fonts/ /usr/share/fonts/truetype
RUN chmod 644 /usr/share/fonts/truetype/*
## Refresh system font cache
RUN apt-get update; apt-get install -y fontconfig
RUN fc-cache -f -vr
RUN fc-list
# update write permissions
RUN chmod 777 .
EXPOSE 80

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY ["AsposeIssueTest.csproj", ""]
RUN dotnet restore "AsposeIssueTest.csproj"
COPY . .
WORKDIR "/src/"
RUN dotnet build "AsposeIssueTest.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "AsposeIssueTest.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /home/site/wwwroot
COPY --from=publish /app/publish .
ENV AzureWebJobsScriptRoot=/home/site/wwwroot \
    AzureFunctionsJobHost__Logging__Console__IsEnabled=true