all : init build read format 

init : init.cpp xlsx2tcpp.hpp
	g++ -std=c++17 -Wall -g -I. init.cpp -lzip -lz --output init

build : build.cpp xlsx2tcpp.hpp
	g++ -std=c++17 -Wall -g -I. -pthread build.cpp -lzip -lz -lstdc++fs --output build

read : read.cpp xlsx2tcpp.hpp
	g++ -std=c++17 -Wall -g -I. -pthread read.cpp -lzip -lz --output read

format :
	clang-format -i xlsx2tcpp.hpp init.cpp build.cpp read.cpp

