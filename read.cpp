#include "test-sheet1.hpp"
#include "xlsx2tcpp.hpp"

int
main()
{

	auto const table{ xlsx2tcpp::read<test_sheet1>() };

	for (auto const& row : table)
		std::cout << row.a << '\n';

	auto const [index, ok]{ xlsx2tcpp::make_index(table, &test_sheet1::a) };

	std::cout << "index.find(2)->second" << '\t' << index.find(2)->second << '\n';

	return 0;
}
