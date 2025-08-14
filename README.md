# Fire Player (.fire)

Небольшое приложение на C++ (GLFW + ImGui + OpenGL), которое загружает файл эффекта с расширением .fire (JSON) и воспроизводит шейдер на полноэкранном квадратике.

## Сборка

Требуется: CMake (>=3.16), компилятор с поддержкой C++20, интернет для загрузки зависимостей (FetchContent).

```bash
mkdir -p build && cd build
cmake .. -DCMAKE_BUILD_TYPE=Release
cmake --build . -j
```

## Запуск

```bash
./fire_player ../assets/examples/fire_glow.fire
```

Если путь к .fire не указан, откроется встроенный пример эффекта.

## Формат .fire (JSON)

Минимальный пример:

```json
{
  "name": "My Effect",
  "width": 1280,
  "height": 720,
  "uniforms": [
    { "name": "uStrength", "type": "float", "value": 1.0 }
  ],
  "vertex": "...GLSL vertex...",
  "fragment": "...GLSL fragment..."
}
```

Поддерживаемые типы uniform: float, int, vec2, vec3, vec4.

Автоматически доступны uniform-переменные:
- iTime: время с запуска (сек)
- iResolution: размер буфера кадра (px)
