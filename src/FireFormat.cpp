#include "FireFormat.h"

#include <fstream>
#include <sstream>

#include <nlohmann/json.hpp>

using nlohmann::json;

static bool readTextFile(const std::string& path, std::string& out) {
	std::ifstream ifs(path, std::ios::in | std::ios::binary);
	if (!ifs) return false;
	std::ostringstream ss;
	ss << ifs.rdbuf();
	out = ss.str();
	return true;
}

static std::string getDirectory(const std::string& path) {
	size_t pos = path.find_last_of("/\\");
	if (pos == std::string::npos) return std::string();
	return path.substr(0, pos);
}

static std::string joinPath(const std::string& dir, const std::string& name) {
	if (dir.empty()) return name;
	char sep = '/';
	if (!dir.empty() && (dir.back() == '/' || dir.back() == '\\')) return dir + name;
	return dir + sep + name;
}

static void fillPassFromJson(const json& jp, const std::string& baseDir, FirePass& outPass) {
	outPass.name = jp.value("name", std::string("pass"));
	std::string vpath = jp.value("vertex_path", std::string());
	std::string fpath = jp.value("fragment_path", std::string());
	if (!vpath.empty()) {
		std::string full = joinPath(baseDir, vpath);
		readTextFile(full, outPass.vertexShader);
	} else {
		outPass.vertexShader = jp.value("vertex", std::string());
	}
	if (!fpath.empty()) {
		std::string full = joinPath(baseDir, fpath);
		readTextFile(full, outPass.fragmentShader);
	} else {
		outPass.fragmentShader = jp.value("fragment", std::string());
	}
}

std::optional<FireEffect> loadFireFromFile(const std::string& filePath, std::string& outError) {
	outError.clear();
	std::string content;
	if (!readTextFile(filePath, content)) {
		outError = "Cannot read file: " + filePath;
		return std::nullopt;
	}

	json j;
	try {
		j = json::parse(content);
	} catch (const std::exception& e) {
		outError = std::string("JSON parse error: ") + e.what();
		return std::nullopt;
	}

	FireEffect effect;
	effect.name = j.value("name", std::string("Unnamed Effect"));
	effect.width = j.value("width", 1280);
	effect.height = j.value("height", 720);

	if (j.contains("uniforms") && j["uniforms"].is_array()) {
		for (const auto& ju : j["uniforms"]) {
			FireUniform u;
			u.name = ju.value("name", std::string());
			u.type = ju.value("type", std::string("float"));
			if (ju.contains("value")) {
				if (ju["value"].is_array()) {
					for (const auto& v : ju["value"]) u.value.push_back(v.get<float>());
				} else if (ju["value"].is_number()) {
					u.value.push_back(ju["value"].get<float>());
				}
			}
			if (!u.name.empty()) effect.uniforms.push_back(std::move(u));
		}
	}

	std::string baseDir = getDirectory(filePath);
	if (j.contains("passes") && j["passes"].is_array()) {
		for (const auto& jp : j["passes"]) {
			FirePass p;
			fillPassFromJson(jp, baseDir, p);
			effect.passes.push_back(std::move(p));
		}
	} else {
		// single-pass convenience keys
		FirePass p;
		fillPassFromJson(j, baseDir, p);
		effect.passes.push_back(std::move(p));
	}

	if (effect.passes.empty()) {
		outError = "No pass found in .fire file";
		return std::nullopt;
	}

	// Basic validation
	if (effect.passes.front().fragmentShader.empty()) {
		outError = "Fragment shader source is empty";
		return std::nullopt;
	}

	return effect;
}