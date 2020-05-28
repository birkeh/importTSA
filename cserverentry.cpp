#include "cserverentry.h"

#include <iostream>
#include "libxl.h"

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
