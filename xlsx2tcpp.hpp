#ifndef XLSX2TCPP_HPP
#define XLSX2TCPP_HPP

#include <fd-read-xlsx.hpp>
#include <filesystem>
#include <fstream>
#include <thread>

#include <cassert>
#include <cstdio>
#include <cstring>
#include <zlib.h>

namespace xlsx2tcpp {

typedef std::string str_t;

class Exception : public std::exception
{

public:
	Exception(str_t const& msg)
	  : msg_("xlsx2tcpp library: " + msg + '.')
	{}
	const char* what() const throw() { return msg_.c_str(); }

private:
	str_t const msg_;
};

std::pair<str_t, str_t>
get_names(char const* const xlsx_file_name, str_t const& sheetname)
{

	str_t const xlsx_file_name_base{ [&]() {
		str_t const name{ xlsx_file_name };
		str_t const ending{ ".xlsx" };
		str_t const ENDING{ ".XLSX" };
		if ((name.size() > ending.size()) &&
		    (std::equal(crbegin(ending), crend(ending), crbegin(name)) ||
		     std::equal(crbegin(ENDING), crend(ENDING), crbegin(name))))
			return str_t(cbegin(name), cend(name) - ending.size());
		return name;
	}() };

	str_t file_name, struct_name;
	for (auto const& c : str_t{ xlsx_file_name_base } + '-' + sheetname) {
		if (std::isalpha(c))
			file_name += c, struct_name += c;
		else if (std::isdigit(c))
			file_name += c, struct_name += struct_name.empty() ? '_' : c;
		else
			file_name += '-', struct_name += '_';
	}
	return { file_name, struct_name };
}

void
init(char const* const xlsx_file_name, char const* const sheet_name = "")
{
	auto const [table, sheetname]{ fd_read_xlsx::get_table_sheetname(xlsx_file_name, sheet_name) };
	if (table.size() < 2)
		throw Exception("the number of rows in the worksheet is less than 2");
	for (size_t j{ 0 }; j < table[0].size(); ++j) {
		if (!fd_read_xlsx::holds_string(table[0][j]))
			throw Exception("a cell of the first row in the worksheet is not a string one");
		if (fd_read_xlsx::get_string(table[0][j]).empty())
			throw Exception("a cell of the first row in the worksheet is empty");
		auto const c{ fd_read_xlsx::get_string(table[0][j])[0] };
		if (!(std::isalpha(c) || (c == '_')))
			throw Exception("a string in a cell of the first row in the worksheet is "
			                "not a valid C++ identifier");
	}
	auto const nr_cols{ table[0].size() };
	for (size_t i{ 1 }; i < table.size(); ++i)
		if (table[i].size() != nr_cols)
			throw Exception("the number of cols is variable between rows");
	std::vector<bool> is_str(nr_cols, false);
	std::vector<size_t> str_szs(nr_cols);
	std::vector<bool> is_int(nr_cols, false);
	for (size_t j{ 0 }; j < nr_cols; ++j) {
		// Test for string.
		bool flag{ true };
		for (size_t i{ 1 }; i < table.size(); ++i)
			if (!fd_read_xlsx::holds_string(table[i][j])) {
				flag = false;
				break;
			}
		if (flag) {
			is_str[j] = true;
			size_t max{ 0 };
			for (size_t i{ 1 }; i < table.size(); ++i) {
				auto const& str{ fd_read_xlsx::get_string(table[i][j]) };
				if (str.size() > max)
					max = str.size();
			}
			str_szs[j] = max;
		} else {
			// Test for int.
			flag = true;
			for (size_t i{ 1 }; i < table.size(); ++i)
				if (!fd_read_xlsx::holds_int(table[i][j])) {
					flag = false;
					break;
				}
			if (flag)
				is_int[j] = true;
			else {
				// Test for int.
				flag = true;
				for (size_t i{ 1 }; i < table.size(); ++i)
					if (!fd_read_xlsx::holds_num(table[i][j])) {
						flag = false;
						break;
					}
				if (!flag)
					throw Exception("the type of the cells is variable between rows");
			}
		}
	}
	auto const [file_name, struct_name]{ get_names(xlsx_file_name, sheetname) };

	std::ofstream out{ file_name + ".hpp" };
	if (!out.is_open())
		throw Exception{ "unable to open for output “" + file_name + ".hpp”" };
	out << "#include <xlsx2tcpp.hpp>\n";
	out << "#include <array>\n\n";
	out << "struct " << struct_name << "{\n";
	// Declare and initialize a static data member with some usefull information.
	out << "\tstruct { size_t n; char const *file_name; char const *struct_name; size_t nr_threads; "
	       "}\n";
	out << "\t\tstatic constexpr _info_ {\n";
	out << "\t\t\t" << (table.size() - 1) << ", \"" << file_name << "\", \"" << struct_name << "\", "
	    << std::thread::hardware_concurrency() << " };\n";
	// The data members.
	for (size_t j{ 0 }; j < nr_cols; ++j) {
		if (is_str[j])
			out << "\tstd::array<char, " << str_szs[j] << "> " << fd_read_xlsx::get_string(table[0][j])
			    << ";\n";
		else if (is_int[j])
			out << "\tint64_t " << fd_read_xlsx::get_string(table[0][j]) << ";\n";
		else
			out << "\tdouble " << fd_read_xlsx::get_string(table[0][j]) << ";\n";
	}
	// The constructors.
	out << '\t' << struct_name << "() {}\n";
	out << '\t' << struct_name << "(std::vector<fd_read_xlsx::cell_t> const & _v_):\n";
	{
		// int64_t and double data members are initialized as data members.
		bool first{ true };
		for (size_t j{ 0 }; j < nr_cols; ++j) {
			if (is_str[j])
				;
			else if (is_int[j]) {
				if (first)
					out << "\t\t  ", first = false;
				else
					out << "\t\t, ";
				out << fd_read_xlsx::get_string(table[0][j]) << "(fd_read_xlsx::get_int(_v_[" << j
				    << "]))\n";
			} else {
				if (first)
					out << "\t\t  ", first = false;
				else
					out << "\t\t, ";
				out << fd_read_xlsx::get_string(table[0][j]) << "(fd_read_xlsx::get_num(_v_[" << j
				    << "]))\n";
			}
		}
	}
	// str data members are initialized within the constructor.
	out << "\t\t{\n";
	for (size_t j{ 0 }; j < nr_cols; ++j) {
		if (is_str[j]) {
			auto const name{ fd_read_xlsx::get_string(table[0][j]) };
			out << "\t\t\t{\n";
			// fill the string with 0 as the defaut initialization leaves the contents of the array
			// indeterminated.
			out << "\t\t\t\t" << name << ".fill('\\0');\n";
			out << "\t\t\t\tauto const str {fd_read_xlsx::get_string(_v_[" << j << "])};\n";
			out << "\t\t\t\tstd::copy(cbegin(str), cend(str), begin(" << name << "));\n";
			out << "\t\t\t}\n";
		}
	}
	out << "\t\t}\n";

	out << "};\n";
}
template<typename T>
void
task_write(str_t const& file_name, std::vector<T> const& tcpp, size_t start, size_t end)
{
	if (start == end)
		return;
	// compression level 9.
	auto out{ gzopen(file_name.c_str(), "wb9") };
	if (out == NULL)
		throw Exception{ "unable to open for output the “" + file_name + "” file" };
	size_t const n =
	  gzwrite(out, reinterpret_cast<char const*>(&tcpp[start]), (end - start) * sizeof(T));
	auto const ret{ gzclose(out) };
	if (n != (end - start) * sizeof(T))
		throw Exception{ "unable to append the zipped stream into the “" + file_name + "” file" };
	if (ret != Z_OK)
		throw Exception{ "unable to close the “" + file_name + "” file" };
}
template<typename T>
void
build(char const* const xlsx_file_name, char const* const sheet_name = "")
{
	auto const table{ fd_read_xlsx::read(xlsx_file_name, sheet_name) };
	if (T::_info_.n != (table[0].size() - 1))
		throw Exception("T::_info_.n != (table[0].size()-1)");
	std::vector<T> tcpp;
	tcpp.reserve(table.size() - 1);
	for (size_t i{ 1 }; i < table.size(); ++i)
		tcpp.push_back(T{ table[i] });
	// Be careful to create the directory.
	if (!std::filesystem::exists(T::_info_.file_name))
		if (!std::filesystem::create_directory(T::_info_.file_name))
			throw Exception{ "unable to create “" + str_t{ T::_info_.file_name } + "” directory" };

	auto const nr_threads{ T::_info_.nr_threads };
	std::vector<std::thread> threads;
	threads.reserve(nr_threads);
	auto const step{ tcpp.size() / nr_threads + 1 };
	size_t start{ 0 };
	for (size_t i = 0; i < nr_threads; ++i) {
		auto const name{ str_t{ T::_info_.file_name } + '/' + std::to_string(i) + ".gz" };
		auto const end{ std::min(start + step, tcpp.size()) };
		threads.emplace_back(task_write<T>, name, tcpp, start, end);
		start = end;
	}
	for (auto& thread : threads)
		thread.join();
}
template<typename T>
void
task_read(str_t const& file_name, std::vector<T>& rvo, size_t start, size_t end)
{
	if (start == end)
		return;
	auto in{ gzopen(file_name.c_str(), "rb") };
	if (in == NULL)
		throw Exception{ "unable to open for input the “" + file_name + "” file" };
	size_t const n = gzread(in, reinterpret_cast<char*>(&rvo[start]), (end - start) * sizeof(T));
	auto const ret{ gzclose(in) };
	if (n != (end - start) * sizeof(T))
		throw Exception{ "unable to read the zipped stream from the “" + file_name + "” file" };
	if (ret != Z_OK)
		throw Exception{ "unable to close the “" + file_name + "” file" };
}
template<typename T>
std::vector<T>
read()
{
	std::vector<T> rvo{ T::_info_.n };
	auto const nr_threads{ T::_info_.nr_threads };
	std::vector<std::thread> threads;
	threads.reserve(nr_threads);
	auto const step{ rvo.size() / nr_threads + 1 };
	size_t start{ 0 };
	for (size_t i = 0; i < nr_threads; ++i) {
		auto const name{ str_t{ T::_info_.file_name } + '/' + std::to_string(i) + ".gz" };
		auto const end{ std::min(start + step, rvo.size()) };
		threads.emplace_back(task_read<T>, name, std::ref(rvo), start, end);
		start = end;
	}
	for (auto& thread : threads)
		thread.join();
	return rvo;
}
} // namespace xlsx2tcpp
#endif // XLSX2TCPP_HPP
