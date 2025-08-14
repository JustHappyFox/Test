#include "FireRuntime.h"

#include <cstdio>
#include <cstdlib>
#include <cmath>
#include <vector>
#include <string>

#include <imgui.h>
#include "imgui_impl_glfw.h"
#include "imgui_impl_opengl3.h"

static void glfwErrorCallback(int error, const char* description) {
	std::fprintf(stderr, "GLFW Error %d: %s\n", error, description);
}

static void glfwDropCallback(GLFWwindow* window, int count, const char** paths) {
	if (count <= 0) return;
	void* ptr = glfwGetWindowUserPointer(window);
	if (!ptr) return;
	FireRuntime* rt = reinterpret_cast<FireRuntime*>(ptr);
	std::string err;
	if (!rt->loadEffectFromPath(paths[0], err)) {
		std::fprintf(stderr, "Failed to load dropped file: %s\n", err.c_str());
	}
}

FireRuntime::FireRuntime()
	: window(nullptr), windowWidth(1280), windowHeight(720), imguiInitialized(false), startTime(0.0) {
}

FireRuntime::~FireRuntime() {
	shutdown();
}

bool FireRuntime::init(const char* title, int width, int height, std::string& outError) {
	outError.clear();
	glfwSetErrorCallback(glfwErrorCallback);
	if (!glfwInit()) {
		outError = "Failed to initialize GLFW";
		return false;
	}

	glfwWindowHint(GLFW_CONTEXT_VERSION_MAJOR, 3);
	glfwWindowHint(GLFW_CONTEXT_VERSION_MINOR, 3);
	glfwWindowHint(GLFW_OPENGL_PROFILE, GLFW_OPENGL_CORE_PROFILE);

	windowWidth = width;
	windowHeight = height;
	window = glfwCreateWindow(width, height, title, nullptr, nullptr);
	if (!window) {
		outError = "Failed to create GLFW window";
		glfwTerminate();
		return false;
	}

	glfwMakeContextCurrent(window);
	glfwSwapInterval(1);

    if (glewInit() != GLEW_OK) {
        outError = "Failed to initialize GLEW";
        return false;
    }

	if (!initImGui()) {
		outError = "Failed to initialize ImGui";
		return false;
	}

	// Setup drag-and-drop
	glfwSetWindowUserPointer(window, this);
	glfwSetDropCallback(window, glfwDropCallback);

	startTime = glfwGetTime();

	return true;
}

bool FireRuntime::initImGui() {
	IMGUI_CHECKVERSION();
	ImGui::CreateContext();
	ImGuiIO& io = ImGui::GetIO();
	(void)io;
	ImGui::StyleColorsDark();
	if (!ImGui_ImplGlfw_InitForOpenGL(window, true)) return false;
	if (!ImGui_ImplOpenGL3_Init("#version 330")) return false;
	imguiInitialized = true;
	return true;
}

void FireRuntime::shutdownImGui() {
	if (!imguiInitialized) return;
	ImGui_ImplOpenGL3_Shutdown();
	ImGui_ImplGlfw_Shutdown();
	ImGui::DestroyContext();
	imguiInitialized = false;
}

void FireRuntime::shutdown() {
	shutdownImGui();
	if (window) {
		glfwDestroyWindow(window);
		window = nullptr;
	}
	glfwTerminate();
}

bool FireRuntime::loadEffect(const FireEffect& effect, std::string& outError) {
	outError.clear();
	currentEffect = effect;

	const FirePass& pass = currentEffect.passes.front();

	std::string log;
	if (!shader.compileFromSource(pass.vertexShader, pass.fragmentShader, log)) {
		outError = log;
		return false;
	}

	// Resize window in case effect defines custom size
	if (currentEffect.width > 0 && currentEffect.height > 0) {
		glfwSetWindowSize(window, currentEffect.width, currentEffect.height);
		windowWidth = currentEffect.width;
		windowHeight = currentEffect.height;
	}

	return true;
}

bool FireRuntime::loadEffectFromPath(const std::string& filePath, std::string& outError) {
	auto loaded = loadFireFromFile(filePath, outError);
	if (!loaded) return false;
	lastLoadedPath = filePath;
	return loadEffect(*loaded, outError);
}

void FireRuntime::renderFrame() {
	int fbWidth, fbHeight;
	glfwGetFramebufferSize(window, &fbWidth, &fbHeight);
	glViewport(0, 0, fbWidth, fbHeight);

	float clearColor[4] = {0.02f, 0.02f, 0.04f, 1.0f};
	glClearColor(clearColor[0], clearColor[1], clearColor[2], clearColor[3]);
	glClear(GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT);

	shader.use();

	double now = glfwGetTime();
	float time = static_cast<float>(now - startTime);
	shader.setUniformFloat("iTime", time);
	shader.setUniformVec2("iResolution", static_cast<float>(fbWidth), static_cast<float>(fbHeight));

	for (const auto& u : currentEffect.uniforms) {
		if (u.type == "float" && u.value.size() >= 1) shader.setUniformFloat(u.name.c_str(), u.value[0]);
		else if (u.type == "int" && u.value.size() >= 1) shader.setUniformInt(u.name.c_str(), static_cast<int>(u.value[0]));
		else if (u.type == "vec2" && u.value.size() >= 2) shader.setUniformVec2(u.name.c_str(), u.value[0], u.value[1]);
		else if (u.type == "vec3" && u.value.size() >= 3) shader.setUniformVec3(u.name.c_str(), u.value[0], u.value[1], u.value[2]);
		else if (u.type == "vec4" && u.value.size() >= 4) shader.setUniformVec4(u.name.c_str(), u.value[0], u.value[1], u.value[2], u.value[3]);
	}

	quad.draw();

	// ImGui overlay
	ImGui_ImplOpenGL3_NewFrame();
	ImGui_ImplGlfw_NewFrame();
	ImGui::NewFrame();

	ImGui::Begin(".fire player");
	ImGui::Text("Effect: %s", currentEffect.name.c_str());
	ImGui::Text("Resolution: %dx%d", fbWidth, fbHeight);
	ImGui::Text("iTime: %.2f s", time);
	if (!lastLoadedPath.empty()) {
		ImGui::TextWrapped("File: %s", lastLoadedPath.c_str());
		if (ImGui::Button("Reload")) {
			std::string err;
			if (!loadEffectFromPath(lastLoadedPath, err)) {
				std::fprintf(stderr, "Reload failed: %s\n", err.c_str());
			}
		}
	}
	ImGui::TextDisabled("Drag & drop a .fire file into the window to load.");
	ImGui::End();

	ImGui::Render();
	ImGui_ImplOpenGL3_RenderDrawData(ImGui::GetDrawData());
}

void FireRuntime::runMainLoop() {
	while (!glfwWindowShouldClose(window)) {
		renderFrame();
		glfwSwapBuffers(window);
		glfwPollEvents();
	}
}