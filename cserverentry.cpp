#include "cserverentry.h"

#include <iostream>
#include "libxl.h"

#include "xlsxdocument.h"

#include <QDir>

#include <QDebug>

using namespace libxl;


cServerEntry::cServerEntry() :
	m_fileName(""),
	m_serverName(""),
	m_reportDate(QDate()),
	m_customerInterface(""),
	m_totalNumberOfUsers(0),
	m_userDisabled(0),
	m_availableDiskSpace(0),
	m_usedDiskSpace(0),
	m_usedDiskSpacePercent(0),
	m_maximumDiskUtil(0),
	m_growthBasePrevMonth(0),
	m_growthBaseStartDate(0),
	m_mainStorage(0),
	m_numberOfProblemRecords(0),
	m_availability(0),
	m_averageCPUUtil(0),
	m_averageCPUUtilTimeframe(""),
	m_maximumCPUUtil(0),
	m_downtimeCum(0)
{
}

bool cServerEntry::load(const QString& fileName)
{
	QString	tmp;
	Book*	book	= xlCreateBook();
	wchar_t	fName[fileName.length()+1];

	fileName.toWCharArray(fName);

	if(book->load(fName))
	{
		Sheet* sheet = book->getSheet(0);
		if(sheet)
		{
			tmp	= QString::fromWCharArray(sheet->readStr(1, 0));

			if(tmp.indexOf(" "))
			{
				tmp	= tmp.mid(tmp.indexOf(" ")+1);

				if(tmp.indexOf(" "))
				{
					m_serverName	= tmp.left(tmp.indexOf(" "));
					tmp				= tmp.mid(tmp.indexOf(" ")+1);
					m_reportDate	= QDate(tmp.left(4).toInt(), tmp.mid(5, 2).toInt(), 1);
				}
			}

			for(int row = 2;row < sheet->lastRow();row++)
			{
				QString	field	= QString::fromWCharArray(sheet->readStr(row, 0));


				if(field == "Customer-Interface")
					m_customerInterface			= QString::fromWCharArray(sheet->readStr(row, 1));
				else if(field == "Total Nbr of User")
					m_totalNumberOfUsers		= static_cast<qint16>(sheet->readNum(row, 1));
				else if(field == "User disabled")
					m_userDisabled				= static_cast<qint16>(sheet->readNum(row, 1));
				else if(field == "Available diskspace (GB)")
					m_availableDiskSpace		= sheet->readNum(row, 1);
				else if(field == "Used diskspace (GB)")
					m_usedDiskSpace				= sheet->readNum(row, 1);
				else if(field == "Used diskspace (%)")
					m_usedDiskSpacePercent		= sheet->readNum(row, 1);
				else if(field == "Maximum disk util.")
					m_maximumDiskUtil			= sheet->readNum(row, 1);
				else if(field == "Growth, base prev. month (%)")
					m_growthBasePrevMonth		= sheet->readNum(row, 1);
				else if(field == "Growth, base startdate  (%)")
					m_growthBaseStartDate		= sheet->readNum(row, 1);
				else if(field == "Mainstorage   (MB)")
					m_mainStorage	= sheet->readNum(row, 1);
				else if(field == "Number of problemrecords")
					m_numberOfProblemRecords	= static_cast<qint16>(sheet->readNum(row, 1));
				else if(field == "Availability    (%)")
					m_availability				= sheet->readNum(row, 1);
				else if(field == "Average CPU util. (%)")
				{
					m_averageCPUUtil			= sheet->readNum(row, 1);
					m_averageCPUUtilTimeframe	= QString::fromWCharArray(sheet->readStr(row, 2));
				}
				else if(field == "Maximum CPU util.  (%)")
					m_maximumCPUUtil			= sheet->readNum(row, 1);
				else if(field == "Downtime cum. (min)")
					m_downtimeCum				= static_cast<qint16>(sheet->readNum(row, 1));
			}
		}
	}

	book->release();

	m_fileName	= fileName;
	return(true);
}

cServerEntryList::cServerEntryList()
{
}

bool cServerEntryList::loadPath(const QString& path)
{
	QDir		dir(path);
	QStringList	fileList	= dir.entryList(QStringList() << "*.xls");

	for(int x = 0;x < fileList.count();x++)
	{
		cServerEntry*	serverEntry	= new cServerEntry;
		if(serverEntry->load(path + "/" + fileList[x]))
			append(serverEntry);
		else
			delete serverEntry;
	}
	return(true);
}

bool cServerEntryList::save(const QString& fileName)
{
	QXlsx::Format		format;
	QXlsx::Format		formatBig;
	QXlsx::Format		formatMerged;
	QXlsx::Format		formatCurrency;
	QXlsx::Format		formatNumber;
	QXlsx::Format		formatDate;
	QXlsx::Format		formatWrap;

	QString				szNumberFormatCurrency("_-\"€\"\\ * #,##0.00_-;\\-\"€\"\\ * #,##0.00_-;_-\"€\"\\ * \"-\"??_-;_-@_-");
	QString				szNumberFormatNumber("#,##0.00");
	QString				szNumberFormatDate("dd.mm.yyyy");

	format.setFontBold(true);

	formatBig.setFontBold(true);
	formatBig.setFontSize(16);

	formatMerged.setHorizontalAlignment(QXlsx::Format::AlignLeft);
	formatMerged.setVerticalAlignment(QXlsx::Format::AlignTop);

	formatCurrency.setNumberFormat(szNumberFormatCurrency);
	formatNumber.setNumberFormat(szNumberFormatNumber);
	formatDate.setNumberFormat(szNumberFormatDate);

	formatWrap.setTextWarp(true);

	QXlsx::Document			xlsx;

	xlsx.write(1,  1, "fileName");
	xlsx.write(1,  2, "serverName");
	xlsx.write(1,  3, "reportDate");
	xlsx.write(1,  4, "customerInterface");
	xlsx.write(1,  5, "totalNumberOfUsers");
	xlsx.write(1,  6, "userDisabled");
	xlsx.write(1,  7, "availableDiskSpace");
	xlsx.write(1,  8, "usedDiskSpace");
	xlsx.write(1,  9, "usedDiskSpacePercent");
	xlsx.write(1, 10, "maximumDiskUtil");
	xlsx.write(1, 11, "growthBasePrevMonth");
	xlsx.write(1, 12, "growthBaseStartDate");
	xlsx.write(1, 13, "mainStorage");
	xlsx.write(1, 14, "numberOfProblemRecords");
	xlsx.write(1, 15, "availability");
	xlsx.write(1, 16, "averageCPUUtil");
	xlsx.write(1, 17, "averageCPUUtilTImeframe");
	xlsx.write(1, 18, "maximumCPUUtil");
	xlsx.write(1, 19, "downtimeCum");

	for(int x = 0;x < count();x++)
	{
		cServerEntry*	serverEntry	= at(x);

		xlsx.write(x+2,  1, serverEntry->m_fileName);
		xlsx.write(x+2,  2, serverEntry->m_serverName);
		xlsx.write(x+2,  3, serverEntry->m_reportDate);
		xlsx.write(x+2,  4, serverEntry->m_customerInterface);
		xlsx.write(x+2,  5, serverEntry->m_totalNumberOfUsers);
		xlsx.write(x+2,  6, serverEntry->m_userDisabled);
		xlsx.write(x+2,  7, serverEntry->m_availableDiskSpace);
		xlsx.write(x+2,  8, serverEntry->m_usedDiskSpace);
		xlsx.write(x+2,  9, serverEntry->m_usedDiskSpacePercent);
		xlsx.write(x+2, 10, serverEntry->m_maximumDiskUtil);
		xlsx.write(x+2, 11, serverEntry->m_growthBasePrevMonth);
		xlsx.write(x+2, 12, serverEntry->m_growthBaseStartDate);
		xlsx.write(x+2, 13, serverEntry->m_mainStorage);
		xlsx.write(x+2, 14, serverEntry->m_numberOfProblemRecords);
		xlsx.write(x+2, 15, serverEntry->m_availability);
		xlsx.write(x+2, 16, serverEntry->m_averageCPUUtil);
		xlsx.write(x+2, 17, serverEntry->m_averageCPUUtilTimeframe);
		xlsx.write(x+2, 18, serverEntry->m_maximumCPUUtil);
		xlsx.write(x+2, 19, serverEntry->m_downtimeCum);
	}

	xlsx.saveAs(fileName);

	return(true);
}
