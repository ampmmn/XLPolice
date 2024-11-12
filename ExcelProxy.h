#pragma once

#include <memory>

class ExcelProxy
{
public:
	ExcelProxy();
	~ExcelProxy();

public:
	bool IsAvailable();
	bool HasDocument();
	void Terminate();

private:
	struct PImpl;
	std::unique_ptr<PImpl> in;
};


