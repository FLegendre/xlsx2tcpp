#ifndef XLSX2TCPP_HPP
#define XLSX2TCPP_HPP

#include <fd-read-xlsx-header-only.hpp>
#include <filesystem>
#include <fstream>
#include <map>
#include <thread>

#include <cassert>
#include <cmath>
#include <cstdio>
#include <cstring>
#include <zlib.h>

// We need to define
// 1) xlsx2tcpp::missing(std::array<char, N>);
// 2) std::to_string(std::array<char, N>) which use the former function.
// So we open the namespace xlsx2tcpp et we close it. Then we open the namespace std and we close
// it.

namespace xlsx2tcpp {
template<size_t N>
bool
missing(std::array<char, N> const& a)
{
	for (size_t i{ 0 }; i < N; ++i)
		if (a[i] != '\0')
			return false;
	return true;
}
} // namespace xlsx2tcpp

// This template is put in the std namespace (bad pratice ?).
namespace std {
template<size_t N>
std::string
to_string(std::array<char, N> const& str)
{
	std::string rvo;
	if (xlsx2tcpp::missing(str))
		rvo += "-.-";
	else {
		for (auto const& c : str)
			if (c)
				rvo += c;
	}
	return rvo;
}
} // namespace std

namespace xlsx2tcpp {

typedef std::string str_t;

class Exception : public std::exception
{

public:
	Exception(str_t const& msg)
	  : msg_("xlsx2tcpp library: " + msg + ((*crbegin(msg) == '?') ? ' ' : '.'))
	{}
	const char* what() const throw() { return msg_.c_str(); }

private:
	str_t const msg_;
};

std::pair<str_t, str_t>
get_names(char const* const xlsx_file_name, str_t const& sheetname)
{
	// Get the base file name et chop the extension if it is an Excel one.
	str_t const xlsx_file_name_base{ [&]() {
		str_t const name{ xlsx_file_name };
		auto const last{ name.rfind('/') };
		auto const begin{ ((last == str_t::npos) && (last != name.size() - 1)) ? 0 : last + 1 };
		str_t const end{ ".xlsx" };
		str_t const END{ ".XLSX" };
		if ((name.size() > end.size()) && (std::equal(crbegin(end), crend(end), crbegin(name)) ||
		                                   std::equal(crbegin(END), crend(END), crbegin(name))))
			return str_t(cbegin(name) + begin, cend(name) - end.size());
		return (begin == 0) ? name : str_t(cbegin(name) + begin, cend(name));
	}() };

	str_t file_name, struct_name;
	// If same names, do not duplicate.
	auto const tmp{ (xlsx_file_name_base == sheetname) ? sheetname
		                                                 : (xlsx_file_name_base + '-' + sheetname) };
	for (auto const& c : tmp) {
		if (std::isalpha(c))
			file_name += c, struct_name += struct_name.empty() ? std::toupper(c) : c;
		else if (std::isdigit(c))
			file_name += c, struct_name += struct_name.empty() ? '_' : c;
		// Dont put a `-' in front of a file name.
		else {
			if (!file_name.empty())
				file_name += '-';
			struct_name += '_';
		}
	}
	return { file_name, struct_name };
}
namespace internals {
void
init(char const* const xlsx_file_name, char const* const sheet_name, bool lower)
{
	std::cout << "Reading “" << xlsx_file_name << "”...\n";
	auto const [table, sheetname]{ fd_read_xlsx::get_table_sheetname(xlsx_file_name, sheet_name) };
	std::cout << "Analysing “" << sheetname << "” sheet of “" << xlsx_file_name << "”...\n";
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
		if (table[i].size() > nr_cols)
			throw Exception("the number of cols is variable between rows at row " +
			                std::to_string(i + 1));
	std::vector<bool> is_str(nr_cols, false);
	std::vector<size_t> str_szs(nr_cols);
	std::vector<bool> is_int(nr_cols, false);
	for (size_t j{ 0 }; j < nr_cols; ++j) {
		std::cout << "Analysing “" << fd_read_xlsx::get_string(table[0][j]) << "” row...\n";
		// Test for string.
		bool flag{ true };
		for (size_t i{ 1 }; i < table.size(); ++i)
			if ((j < table[i].size()) && !fd_read_xlsx::holds_string(table[i][j])) {
				flag = false;
				break;
			}
		if (flag) {
			is_str[j] = true;
			size_t max{ 0 };
			for (size_t i{ 1 }; i < table.size(); ++i) {
				if (j < table[i].size()) {
					auto const& str{ fd_read_xlsx::get_string(table[i][j]) };
					if (str.size() > max)
						max = str.size();
				}
			}
			str_szs[j] = max;
		} else {
			// Test for int.
			flag = true;
			for (size_t i{ 1 }; i < table.size(); ++i)
				if ((j < table[i].size()) && !fd_read_xlsx::holds_int(table[i][j]) &&
				    !fd_read_xlsx::empty(table[i][j])) {
					flag = false;
					break;
				}
			if (flag)
				is_int[j] = true;
			else {
				// Test for double.
				flag = true;
				for (size_t i{ 1 }; i < table.size(); ++i)
					if ((j < table[i].size()) && !fd_read_xlsx::holds_num(table[i][j]) &&
					    !fd_read_xlsx::empty(table[i][j])) {
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
	// Convert the name of the variables to lowercase.
	auto const to_lower{ [&](str_t const& str) {
		if (!lower)
			return str;
		str_t rvo;
		std::transform(
		  begin(str), end(str), back_inserter(rvo), [](auto c) { return std::tolower(c); });
		return rvo;
	} };
	// The data members.
	for (size_t j{ 0 }; j < nr_cols; ++j) {
		if (is_str[j])
			out << "\tstd::array<char, " << str_szs[j] << "> "
			    << to_lower(fd_read_xlsx::get_string(table[0][j])) << ";\n";
		else if (is_int[j])
			out << "\tint64_t " << to_lower(fd_read_xlsx::get_string(table[0][j])) << ";\n";
		else
			out << "\tdouble " << to_lower(fd_read_xlsx::get_string(table[0][j])) << ";\n";
	}
	// The constructors.
	out << '\t' << struct_name << "() {}\n";
	out << '\t' << struct_name << "(std::vector<fd_read_xlsx::cell_t> const & _v_)\n";
	{
		// int64_t and double data members are initialized as data members.
		bool first{ true };
		for (size_t j{ 0 }; j < nr_cols; ++j) {
			if (is_str[j])
				;
			else if (is_int[j]) {
				if (first)
					out << "\t\t: ", first = false;
				else
					out << "\t\t, ";
				out << to_lower(fd_read_xlsx::get_string(table[0][j])) << "(((" << j
				    << " < _v_.size()) && !fd_read_xlsx::empty(_v_[" << j
				    << "])) ? fd_read_xlsx::get_int(_v_[" << j
				    << "]) : std::numeric_limits<int64_t>::max())\n";
			} else {
				if (first)
					out << "\t\t  ", first = false;
				else
					out << "\t\t, ";
				out << to_lower(fd_read_xlsx::get_string(table[0][j])) << "(((" << j
				    << " < _v_.size()) && !fd_read_xlsx::empty(_v_[" << j
				    << "])) ? fd_read_xlsx::get_num(_v_[" << j
				    << "]) : std::numeric_limits<double>::quiet_NaN())\n";
			}
		}
	}
	// str data members are initialized within the constructor.
	out << "\t\t{\n";
	for (size_t j{ 0 }; j < nr_cols; ++j) {
		if (is_str[j]) {
			auto const name{ to_lower(fd_read_xlsx::get_string(table[0][j])) };
			out << "\t\t\t{\n";
			// fill the string with 0 as the defaut initialization leaves the contents of the array
			// indeterminated.
			out << "\t\t\t\t" << name << ".fill('\\0');\n";
			out << "\t\t\t\tif ( " << j << " < _v_.size() ) {\n";
			out << "\t\t\t\t\tauto const str {fd_read_xlsx::get_string(_v_[" << j << "])};\n";
			out << "\t\t\t\t\tstd::copy(cbegin(str), cend(str), begin(" << name << "));\n";
			out << "\t\t\t\t}\n";
			out << "\t\t\t}\n";
		}
	}
	out << "\t\t}\n";

	out << "};\n";
}
}
void
init(char const* const xlsx_file_name, char const* const sheet_name = "")
{
	internals::init(xlsx_file_name, sheet_name, false);
}
void
lower_init(char const* const xlsx_file_name, char const* const sheet_name = "")
{
	internals::init(xlsx_file_name, sheet_name, true);
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
	std::cout << "Reading “" << xlsx_file_name << "”...\n";
	auto const table{ fd_read_xlsx::read(xlsx_file_name, sheet_name) };
	if (T::_info_.n != (table.size() - 1))
		throw Exception("T::_info_.n (" + std::to_string(T::_info_.n) + ") != (table.size()-1) (" +
		                std::to_string(table.size() - 1) + ')');
	std::vector<T> tcpp;
	tcpp.reserve(table.size() - 1);
	std::cout << "Copying “" << T::_info_.file_name << "”...\n";
	for (size_t i{ 1 }; i < table.size(); ++i)
		tcpp.push_back(T{ table[i] });

	std::cout << "Zipping “" << T::_info_.file_name << "”...\n";
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
read(str_t const& dir_name = "")
{
	std::vector<T> rvo{ T::_info_.n };
	auto const nr_threads{ T::_info_.nr_threads };
	std::vector<std::thread> threads;
	threads.reserve(nr_threads);
	auto const step{ rvo.size() / nr_threads + 1 };
	size_t start{ 0 };
	for (size_t i = 0; i < nr_threads; ++i) {
		auto const name{ (dir_name.empty() ? str_t{} : (dir_name + '/')) +
			               str_t{ T::_info_.file_name } + '/' + std::to_string(i) + ".gz" };
		auto const end{ std::min(start + step, rvo.size()) };
		threads.emplace_back(task_read<T>, name, std::ref(rvo), start, end);
		start = end;
	}
	for (auto& thread : threads)
		thread.join();
	return rvo;
}
// Usage
//   for ( auto const & row : table ) {
//       if ( first(table, row.a) ) {
//           ...
//
// It is presumed that table is sorted by the values of a. The function returns
// true if it is the first row or if the value of a from the previous row is
// not equal to the value of current row.
template<typename T, typename U>
bool
first(std::vector<T> const& table, U const& u)
{
	auto const address_table{ reinterpret_cast<char const*>(&table[0]) };
	auto const address_u{ reinterpret_cast<char const*>(&u) };
	assert(address_table <= address_u);
	// Narrowing initializations for the compiler...
	size_t i = (address_u - address_table) / sizeof(T);
	size_t offset = address_u - (address_table + i * sizeof(T));
	if (i == 0)
		return true;
	return std::memcmp(address_table + (i - 1) * sizeof(T) + offset,
	                   address_table + i * sizeof(T) + offset,
	                   sizeof(U)) != 0;
}
template<typename T, typename U>
bool
last(std::vector<T> const& table, U const& u)
{
	auto const address_table{ reinterpret_cast<char const*>(&table[0]) };
	auto const address_u{ reinterpret_cast<char const*>(&u) };
	assert(address_table <= address_u);
	auto const i{ size_t(address_u - address_table) / sizeof(T) };
	auto const offset{ size_t(address_u - (address_table + i * sizeof(T))) };
	if (i >= table.size())
		return true;
	return std::memcmp(address_table + i * sizeof(T) + offset,
	                   address_table + (i + 1) * sizeof(T) + offset,
	                   sizeof(U)) != 0;
}
bool
missing(int64_t i)
{
	return i == std::numeric_limits<int64_t>::max();
}
bool
missing(double x)
{
	return std::isnan(x);
}
// auto const N { not_missing(table, &Row::member) };
template<typename T, typename U>
size_t
not_missing(std::vector<T> const& table, U T::*m_ptr)
{
	size_t N{ 0 };
	for (auto const& row : table)
		if (!missing(row.*m_ptr))
			++N;
	return N;
}
template<typename T>
size_t
num_obs(std::vector<T> const& table, T const& row)
{
	assert(&row >= &table[0]);
	return &row - &table[0];
}
// index(table, &Row::member, key)
template<typename T, typename U>
size_t
index(std::vector<T> const& table, U T::*m_ptr, U const& key)
{
	// The maps are stored in a static main map. In case of miss, the map is created on the fly and
	// putted in the cache. Otherwise, the map in cache is used. The key of the main map is the adress
	// of the first observation of the table and the offset of the data member.
	static std::map<std::pair<T const*, size_t>, std::map<U, size_t>> maps;
	auto const offset{ size_t(reinterpret_cast<char const*>(&(table[0].*m_ptr)) -
		                        reinterpret_cast<char const*>(&table[0])) };
	auto const map_key{ std::make_pair(&table[0], offset) };
	auto const it{ maps.find(map_key) };
	if (it == std::cend(maps)) {
		std::map<U, size_t> tmp;
		for (auto const& row : table)
			tmp[row.*m_ptr] = &row - &table[0];
		maps[map_key] = tmp;
		auto const it_key{ tmp.find(key) };
		if (it_key == std::cend(tmp))
			throw Exception{ "key “" + std::to_string(key) + "“ not found" };
		return it_key->second;
	}
	auto const it_key{ it->second.find(key) };
	if (it_key == std::cend(it->second))
		throw Exception{ "key “" + std::to_string(key) + "“ not found" };
	return it_key->second;
}
//  xt::index(table, &Row::member, key, &Row::get_member)
template<typename T, typename U, typename V>
V
index(std::vector<T> const& table, U T::*m_ptr, U const& key, V T::*m_get_ptr)
{
	return table[index(table, m_ptr, key)].*m_get_ptr;
}
// std::cout << freq(table, &Row::member, name);
template<typename T, typename U>
std::string
freq(std::vector<T> const& table, U T::*m_ptr, std::string const& name)
{
	std::string rvo{ "Freq of " + name + ".\n" };
	if (table.empty())
		rvo += "<empty table>\n";
	else {
		std::map<U, size_t> freq;
		for (auto const& row : table)
			++freq[row.*m_ptr];
		for (auto const& p : freq) {
			auto const pct{ size_t(100. * p.second / table.size()) };
			rvo += (missing(p.first) ? std::string("-.-") : std::to_string(p.first)) + '\t' +
			       std::to_string(p.second) + '\t' + (pct ? std::to_string(pct) : std::string("𝜀")) +
			       '\n';
		}
	}
	return rvo;
}
} // namespace xlsx2tcpp

// This template is put in the global namespace (bad pratice ?).
template<size_t N>
std::ostream&
operator<<(std::ostream& os, std::array<char, N> const& str)
{
	if (xlsx2tcpp::missing(str))
		os << "-.-";
	else {
		for (auto const& c : str)
			if (c)
				os << c;
	}
	return os;
}
#endif // XLSX2TCPP_HPP
