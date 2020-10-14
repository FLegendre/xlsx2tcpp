#include "test-sheet1.hpp"
#include "xlsx2tcpp.hpp"

int
main()
{

	auto const table{ xlsx2tcpp::read<test_sheet1>() };

	for (auto const& row : table)
		std::cout << row.b << '\n';

	return 0;
}
