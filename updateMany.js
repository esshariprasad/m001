db.movieDetails.updateMany({
	rated:null
},{
	$unset:{
	rated:""
	}
}
)