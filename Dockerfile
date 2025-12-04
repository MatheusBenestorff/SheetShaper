#Build
FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build
WORKDIR /src
COPY ["src/SheetShaper.Core/SheetShaper.Core.csproj", "SheetShaper.Core/"]
COPY ["src/SheetShaper.CLI/SheetShaper.CLI.csproj", "SheetShaper.CLI/"]
RUN dotnet restore "SheetShaper.CLI/SheetShaper.CLI.csproj"

COPY src/ .
WORKDIR /src/SheetShaper.CLI
RUN dotnet publish -c Release -o /app/publish

#Run
FROM mcr.microsoft.com/dotnet/runtime:9.0 AS final
WORKDIR /app
COPY --from=build /app/publish .

VOLUME /data

ENTRYPOINT ["dotnet", "SheetShaper.CLI.dll"]