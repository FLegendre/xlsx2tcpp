#include "test-sheet1.hpp"
#include "xlsx2tcpp.hpp"

int
main()
{

	auto const table{ xlsx2tcpp::read<test_sheet1>() };

	for (auto const& row : table)
		std::cout << row.a << '\n';

	std::cout << xlsx2tcpp::index(table, &test_sheet1::a, 2l) << '\n';

	return 0;
}
