#include "test-sheet1.hpp"
#include "xlsx2tcpp.hpp"

int
main()
{

	xlsx2tcpp::build<test_sheet1>("test.xlsx");

	return 0;
}
